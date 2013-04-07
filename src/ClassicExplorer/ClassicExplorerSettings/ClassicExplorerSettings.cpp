// Classic Shell (c) 2009-2013, Ivo Beltchev
// The sources for Classic Shell are distributed under the MIT open source license

#include <windows.h>
#include <commctrl.h>
#include <shlwapi.h>
#include <Psapi.h>
#include "StringUtils.h"

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
			if (_wcsicmp(PathFindFileName(path),L"ClassicExplorerSettings.exe")==0)
			{
				SetForegroundWindow(hwnd);
				bFound=true;
			}
		}
		CloseHandle(hProcess);
	}
	return !bFound;
}

// A simple program that loads ClassicExplorer32.dll and calls the ShowExplorerSettings function
// Why not use rundll32 instead? Because it doesn't include the correct manifest for comctl32.dll
int WINAPI wWinMain( HINSTANCE hInstance, HINSTANCE hPrevInstance, LPTSTR lpstrCmdLine, int nCmdShow )
{
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
	Sprintf(mutexName,_countof(mutexName),L"ClassicExplorerSettings.Mutex.%s.%s",userName,deskName);
	free(deskName);

	HANDLE hMutex=CreateMutex(NULL,TRUE,mutexName);
	if (GetLastError()==ERROR_ALREADY_EXISTS || GetLastError()==ERROR_ACCESS_DENIED)
	{
		EnumWindows(FindSettingsEnum,0);
		return 0;
	}

	wchar_t path[_MAX_PATH];
	GetModuleFileName(NULL,path,_countof(path));
	*PathFindFileName(path)=0;
	wcscat_s(path,L"ClassicExplorer32.dll");

	HMODULE dll=LoadLibrary(path);
	if (!dll) return 1;

	FARPROC proc=GetProcAddress(dll,"ShowExplorerSettings");
	if (!proc) return 2;

	proc();
	return 0;
}
