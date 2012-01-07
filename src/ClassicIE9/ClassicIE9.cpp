// Classic Shell (c) 2009-2011, Ivo Beltchev
// The sources for Classic Shell are distributed under the MIT open source license

#include <windows.h>
#include <commctrl.h>
#include <shlwapi.h>
#include <Psapi.h>
#include "StringUtils.h"
#include "ClassicIE9DLL\ClassicIE9DLL.h"

// Manifest to enable the 6.0 common controls
#pragma comment(linker, \
	"\"/manifestdependency:type='Win32' "\
	"name='Microsoft.Windows.Common-Controls' "\
	"version='6.0.0.0' "\
	"processorArchitecture='*' "\
	"publicKeyToken='6595b64144ccf1df' "\
	"language='*'\"")

// Find and activate the Settings window
static BOOL CALLBACK FindSettingsEnum( HWND hwnd, LPARAM lParam )
{
	wchar_t className[256];
	if (!GetClassName(hwnd,className,_countof(className)) || _wcsicmp(className,L"#32770")!=0)
		return TRUE;
	DWORD process=0;
	GetWindowThreadProcessId(hwnd,&process);
	HANDLE hProcess=OpenProcess(PROCESS_QUERY_INFORMATION|PROCESS_VM_READ,FALSE,process);
	bool bFound=false;
	if (hProcess!=INVALID_HANDLE_VALUE)
	{
		wchar_t path[_MAX_PATH];
		if (GetModuleFileNameEx(hProcess,NULL,path,_countof(path)))
		{
			if (_wcsicmp(PathFindFileName(path),L"ClassicIE9_32.exe")==0)
			{
				SetForegroundWindow(hwnd);
				bFound=true;
			}
		}
		CloseHandle(hProcess);
	}
	return !bFound;
}

void ZoneConfigure( HWND hWnd, const wchar_t *url )
{
	// use undocumented function 383 from shlwapi
	typedef void (WINAPI* FZoneConfigureW)(HWND,LPCWSTR);
	FZoneConfigureW ZoneConfigureW;

	HMODULE	hShlwapi=LoadLibrary(L"shlwapi.dll");
	if(hShlwapi)
	{
		ZoneConfigureW=(FZoneConfigureW)GetProcAddress(hShlwapi,MAKEINTRESOURCEA(383));
		if(ZoneConfigureW)
			ZoneConfigureW(hWnd,url);
		FreeLibrary(hShlwapi);
	}
}

int WINAPI wWinMain( HINSTANCE hInstance, HINSTANCE hPrevInstance, LPTSTR lpCmdLine, int nCmdShow )
{
	if (wcsncmp(lpCmdLine,L"zone ",5)==0)
	{
		wchar_t token[100];
		const wchar_t *url=GetToken(lpCmdLine+5,token,_countof(token),L" ");
		ZoneConfigure((HWND)_wtol(token),url);
		return 0;
	}

	HWND topWindow=(HWND)_wtol(lpCmdLine);
	if (topWindow)
	{
#ifdef _WIN64
		HMODULE hHookModule=GetModuleHandle(L"ClassicIE9DLL_64.dll");
#else
		HMODULE hHookModule=GetModuleHandle(L"ClassicIE9DLL_32.dll");
#endif

		HWND caption=FindWindowEx(topWindow,NULL,L"Client Caption",NULL);
		LogMessage("exe: topWindow=%X, caption=%X\r\n",(DWORD)topWindow,(DWORD)caption);
		UINT message=RegisterWindowMessage(L"ClassicIE9.Injected");
		if (caption)
		{
			if (SendMessage(caption,message,0,0)!=0)
				return 0;

			DWORD processId;
			DWORD threadId=GetWindowThreadProcessId(topWindow,&processId);
			{
				HANDLE hToken;
				if (OpenProcessToken(GetCurrentProcess(),TOKEN_ADJUST_PRIVILEGES|TOKEN_QUERY,&hToken))
				{
					TOKEN_PRIVILEGES tp={1};
					if (LookupPrivilegeValue(NULL,L"SeDebugPrivilege",&tp.Privileges[0].Luid))
						tp.Privileges[0].Attributes=SE_PRIVILEGE_ENABLED;
					AdjustTokenPrivileges(hToken,FALSE,&tp,sizeof(TOKEN_PRIVILEGES),NULL,NULL); 
					CloseHandle(hToken);
				}
			}

			HANDLE hProcess=OpenProcess(PROCESS_ALL_ACCESS,FALSE,processId);
			if (hProcess)
			{
				wchar_t path[_MAX_PATH];
				GetModuleFileName(hHookModule,path,_countof(path));
				void *remotePath=VirtualAllocEx(hProcess,NULL,sizeof(path),MEM_COMMIT,PAGE_READWRITE);
				if (remotePath)
				{
					if (WriteProcessMemory(hProcess,remotePath,path,sizeof(path),NULL))
					{
						HANDLE hThread=CreateRemoteThread(hProcess,NULL,0,(LPTHREAD_START_ROUTINE)GetProcAddress(GetModuleHandle(L"kernel32.dll"),"LoadLibraryW"),remotePath,0,NULL);
						if (hThread)
						{
							WaitForSingleObject(hThread,INFINITE);
							CloseHandle(hThread);
						}
					}
					VirtualFreeEx(hProcess,remotePath,sizeof(path),MEM_RELEASE);
				}
				CloseHandle(hProcess);
			}
		}
		return 0;
	}

#ifndef _WIN64
	if (*lpCmdLine)
#endif
		return 0;

	// if 32-bit exe is called with no arguments, show the settings

	INITCOMMONCONTROLSEX init={sizeof(init),ICC_STANDARD_CLASSES};
	InitCommonControlsEx(&init);
	SetProcessDPIAware();

	// prevent multiple instances from running on the same desktop
	// the assumption is that multiple desktops for the same user will have different name (but may repeat across users)
	wchar_t userName[256];
	DWORD len=_countof(userName);
	GetUserName(userName,&len);
	len=0;
	HANDLE desktop=GetThreadDesktop(GetCurrentThreadId());
	GetUserObjectInformation(desktop,UOI_NAME,NULL,0,&len);
	wchar_t *deskName=(wchar_t*)malloc(len);
	GetUserObjectInformation(desktop,UOI_NAME,deskName,len,&len);

	wchar_t mutexName[1024];
	Sprintf(mutexName,_countof(mutexName),L"ClassicIE9Settings.Mutex.%s.%s",userName,deskName);
	free(deskName);

	HANDLE hMutex=CreateMutex(NULL,TRUE,mutexName);
	if (GetLastError()==ERROR_ALREADY_EXISTS || GetLastError()==ERROR_ACCESS_DENIED)
	{
		EnumWindows(FindSettingsEnum,0);
		return 0;
	}

	ShowIE9Settings();
	return 0;
}
