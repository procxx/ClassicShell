// Classic Shell (c) 2009, Ivo Beltchev
// The sources for Classic Shell are distributed under the MIT open source license

#define _WIN32_WINNT 0x0600
#include <windows.h>
#include <shlwapi.h>
#include <stdio.h>
#include <commctrl.h>
#include "resource.h"

// ClassicShellSetup.exe is a bootstrap application that contains installers for 32-bit and 64-bit.
// It unpacks the right installer into the temp directory and executes it.
// Finally, if the start menu is installed we launch it for the first time. Note: The installer can't
// launch the start menu itself because it runs as the SYSTEM user and we need to run as the logged in user

typedef BOOL (WINAPI *FIsWow64Process)( HANDLE hProcess, PBOOL Wow64Process );

int APIENTRY wWinMain( HINSTANCE hInstance, HINSTANCE hPrevInstance, LPTSTR lpCmdLine, int nCmdShow )
{
	INITCOMMONCONTROLSEX init={sizeof(init),ICC_STANDARD_CLASSES};
	InitCommonControlsEx(&init);
	// check Windows version
	if ((GetVersion()&255)<6)
	{
		MessageBox(NULL,L"Classic Shell requires Windows Vista or later.",L"Classic Shell Setup",MB_OK|MB_ICONERROR);
		return 1;
	}

	// dynamically link to IsWow64Process because it is not available for Windows 2000
	HMODULE hKernel32=GetModuleHandle(L"kernel32.dll");
	FIsWow64Process isWow64Process=(FIsWow64Process)GetProcAddress(hKernel32,"IsWow64Process");
	if (!isWow64Process)
	{
		MessageBox(NULL,L"Classic Shell requires Windows Vista or later.",L"Classic Shell Setup",MB_OK|MB_ICONERROR);
		return 1;
	}

	// extract the installer
	BOOL b64=FALSE;
	isWow64Process(GetCurrentProcess(),&b64);
	void *pRes=NULL;
	HRSRC hResInfo=FindResource(hInstance,MAKEINTRESOURCE(b64?IDR_MSI_FILE64:IDR_MSI_FILE32),L"MSI_FILE");
	if (hResInfo)
	{
		HGLOBAL hRes=LoadResource(hInstance,hResInfo);
		pRes=LockResource(hRes);
	}
	if (!pRes)
	{
		MessageBox(NULL,L"Internal Setup Error",L"Classic Shell Setup",MB_OK|MB_ICONERROR);
		return 1;
	}
	wchar_t path[_MAX_PATH];
	GetTempPath(_countof(path),path);
	wchar_t msiName[_MAX_PATH];
	GetTempFileName(path,L"CSH",0,msiName);
	HANDLE hFile=CreateFile(msiName,GENERIC_WRITE,0,NULL,CREATE_ALWAYS,FILE_ATTRIBUTE_NORMAL,NULL);
	if (hFile==INVALID_HANDLE_VALUE)
	{
		wchar_t message[1024];
		swprintf_s(message,L"Failed to create temp file '%s'.",msiName);
		MessageBox(NULL,message,L"Classic Shell Setup",MB_OK|MB_ICONERROR);
		return 2;
	}
	DWORD q;
	WriteFile(hFile,pRes,SizeofResource(hInstance,hResInfo),&q,NULL);
	CloseHandle(hFile);

	// start the installer
	STARTUPINFO startupInfo;
	memset(&startupInfo,0,sizeof(startupInfo));
	startupInfo.cb=sizeof(startupInfo);
	PROCESS_INFORMATION processInfo;
	memset(&processInfo,0,sizeof(processInfo));
	wchar_t cmdLine[1024];
	swprintf_s(cmdLine,L"msiexec.exe /i %s",msiName);

	if (!CreateProcess(NULL,cmdLine,NULL,NULL,TRUE,0,NULL,NULL,&startupInfo,&processInfo))
	{
		DeleteFile(msiName);
		MessageBox(NULL,L"Failed to run msiexec.exe",L"Classic Shell Setup",MB_OK|MB_ICONERROR);
	}
	else
	{
		// wait for the installer to finish
		WaitForSingleObject(processInfo.hProcess,INFINITE);
		DeleteFile(msiName);
		DWORD code;
		GetExitCodeProcess(processInfo.hProcess,&code);
		if (code) return code;

		// if there were no errors, launch the start menu for the first time
		HKEY hKey=NULL;
		if (RegCreateKeyEx(HKEY_LOCAL_MACHINE,L"SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run",0,NULL,REG_OPTION_NON_VOLATILE,KEY_READ|KEY_QUERY_VALUE|KEY_WOW64_64KEY,NULL,&hKey,NULL)==ERROR_SUCCESS)
		{
			DWORD type=0;
			DWORD size=sizeof(path);
			if (RegQueryValueEx(hKey,L"Classic Start Menu",0,&type,(BYTE*)path,&size)==ERROR_SUCCESS && type==REG_SZ)
			{
				memset(&startupInfo,0,sizeof(startupInfo));
				startupInfo.cb=sizeof(startupInfo);
				memset(&processInfo,0,sizeof(processInfo));
				CreateProcess(NULL,path,NULL,NULL,TRUE,0,NULL,NULL,&startupInfo,&processInfo);
			}
			RegCloseKey(hKey);
		}
	}
	return 0;
}
