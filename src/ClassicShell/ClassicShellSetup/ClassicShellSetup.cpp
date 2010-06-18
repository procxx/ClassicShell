// Classic Shell (c) 2009-2010, Ivo Beltchev
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
	bool bQuiet=wcsstr(lpCmdLine,L"/qn")!=NULL;
	INITCOMMONCONTROLSEX init={sizeof(init),ICC_STANDARD_CLASSES};
	InitCommonControlsEx(&init);
	// check Windows version
	if ((GetVersion()&255)<6)
	{
		if (!bQuiet)
			MessageBox(NULL,L"Classic Shell requires Windows Vista or later.",L"Classic Shell Setup",MB_OK|MB_ICONERROR);
		return 101;
	}

	// dynamically link to IsWow64Process because it is not available for Windows 2000
	HMODULE hKernel32=GetModuleHandle(L"kernel32.dll");
	FIsWow64Process isWow64Process=(FIsWow64Process)GetProcAddress(hKernel32,"IsWow64Process");
	if (!isWow64Process)
	{
		if (!bQuiet)
			MessageBox(NULL,L"Classic Shell requires Windows Vista or later.",L"Classic Shell Setup",MB_OK|MB_ICONERROR);
		return 101;
	}

	BOOL b64=FALSE;
	isWow64Process(GetCurrentProcess(),&b64);

	// check for versions older than 1.0.0
	const wchar_t *oldVersions32[]={
		L"{4FB649CF-3B19-44C2-AE13-3978BA10E3C0}", // 0.9.7
		L"{131E8BB5-6E2F-437B-9923-3BAC5402995D}", // 0.9.8
		L"{962C0EF9-28A6-48B5-AE5D-F8F8B4B1C5F6}", // 0.9.9
		L"{AA86C803-F195-4593-A9EC-24D26D4F9C7E}", // 0.9.10
		NULL
	};

	const wchar_t *oldVersions64[]={
		L"{962E3DB4-82A7-4B38-80B4-F3DB790D9CA2}", // 0.9.7
		L"{4F5A8EAD-D866-47CB-85C3-E17BB328687E}", // 0.9.8
		L"{029C99FA-B112-486A-8350-DA2099C812ED}", // 0.9.9
		L"{2099745F-EFD7-43C8-9A3A-5EAF01CD56FF}", // 0.9.10
		NULL
	};

	
	for (const wchar_t **oldVersion=b64?oldVersions64:oldVersions32;*oldVersion;oldVersion++)
	{
		wchar_t buf[256];
		swprintf_s(buf,L"SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\%s",*oldVersion);
		HKEY hKey=NULL;
		if (RegCreateKeyEx(HKEY_LOCAL_MACHINE,buf,0,NULL,REG_OPTION_NON_VOLATILE,KEY_READ|KEY_QUERY_VALUE|KEY_WOW64_64KEY,NULL,&hKey,NULL)==ERROR_SUCCESS)
		{
			DWORD version;
			DWORD size=sizeof(version);
			if (RegQueryValueEx(hKey,L"Version",0,NULL,(BYTE*)&version,&size)==ERROR_SUCCESS)
			{
				RegCloseKey(hKey);
				if (!bQuiet)
					MessageBox(NULL,L"This version of Classic Shell cannot be installed over versions older than 1.0.0. Please uninstall the old version of Classic Shell, log off, and run the installer again.",L"Classic Shell Setup",MB_OK|MB_ICONERROR);
				return 102;
			}
			RegCloseKey(hKey);
		}
	}
/*
	// warning about being beta
	if (!bQuiet)
	{
		if (MessageBox(NULL,L"Warning!\nThis is a beta version of Classic Shell. It contains features that are not fully tested. Please report any problems in the Source Forge forum. If you prefer a stable build over the latest features, you can download one of the \"general release\" versions like 1.0.3.\nDo you want to continue with the installation?",L"Classic Shell Setup",MB_YESNO|MB_ICONWARNING)==IDNO)
			return 99;
	}
*/

	// the 64-bit version of Classic Shell 1.9.7 has a bug in the uninstaller that fails to back up the ini files. if that version is detected,
	// warn the user to skip the backup step.
	if (b64)
	{
		HKEY hKey;
		if (RegOpenKeyEx(HKEY_LOCAL_MACHINE,L"SOFTWARE\\IvoSoft\\ClassicShell",0,KEY_READ|KEY_WOW64_64KEY,&hKey)==ERROR_SUCCESS)
		{
			DWORD version;
			DWORD size=sizeof(version);
			if (RegQueryValueEx(hKey,L"Version",0,NULL,(BYTE*)&version,&size)==ERROR_SUCCESS && version==10907)
			{
				RegCloseKey(hKey);
				MessageBox(NULL,L"Warning!\nYou are about to upgrade from version 1.9.7 of Classic Shell. There is a known problem with that version on 64-bit systems. When asked if you want to back up the ini files, respond with 'No'. Otherwise the installation will abort.",L"Classic Shell Setup",MB_OK|MB_ICONWARNING);
			}
		}
	}


	// extract the installer
	void *pRes=NULL;
	HRSRC hResInfo=FindResource(hInstance,MAKEINTRESOURCE(b64?IDR_MSI_FILE64:IDR_MSI_FILE32),L"MSI_FILE");
	if (hResInfo)
	{
		HGLOBAL hRes=LoadResource(hInstance,hResInfo);
		pRes=LockResource(hRes);
	}
	if (!pRes)
	{
		if (!bQuiet)
			MessageBox(NULL,L"Internal Setup Error",L"Classic Shell Setup",MB_OK|MB_ICONERROR);
		return 103;
	}
	wchar_t path[_MAX_PATH*2];
	GetTempPath(_countof(path),path);
	wchar_t msiName[_MAX_PATH];
	GetTempFileName(path,L"CSH",0,msiName);
	HANDLE hFile=CreateFile(msiName,GENERIC_WRITE,0,NULL,CREATE_ALWAYS,FILE_ATTRIBUTE_NORMAL,NULL);
	if (hFile==INVALID_HANDLE_VALUE)
	{
		wchar_t message[1024];
		swprintf_s(message,L"Failed to create temp file '%s'.",msiName);
		if (!bQuiet)
			MessageBox(NULL,message,L"Classic Shell Setup",MB_OK|MB_ICONERROR);
		return 104;
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
	wchar_t cmdLine[2048];
	swprintf_s(cmdLine,L"msiexec.exe /i %s %s",msiName,lpCmdLine);

	if (!CreateProcess(NULL,cmdLine,NULL,NULL,TRUE,0,NULL,NULL,&startupInfo,&processInfo))
	{
		DeleteFile(msiName);
		if (!bQuiet)
			MessageBox(NULL,L"Failed to run msiexec.exe",L"Classic Shell Setup",MB_OK|MB_ICONERROR);
		return 105;
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
				if (!bQuiet)
					wcscat_s(path,L" -open");
				CreateProcess(NULL,path,NULL,NULL,TRUE,0,NULL,NULL,&startupInfo,&processInfo);
			}
			RegCloseKey(hKey);
		}
	}
	return 0;
}
