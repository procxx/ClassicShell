// Classic Shell (c) 2009-2010, Ivo Beltchev
// The sources for Classic Shell are distributed under the MIT open source license

// dllmain.cpp : Implementation of DllMain.

#include "stdafx.h"
#include "resource.h"
#include "ClassicExplorer_i.h"
#include "dllmain.h"
#include "..\LocalizationSettings\ParseSettings.h"

CClassicExplorerModule _AtlModule;

void InitClassicCopyProcess( void );
void InitClassicCopyThread( void );
void FreeClassicCopyThread( void );

// DLL Entry Point
extern "C" BOOL WINAPI DllMain( HINSTANCE hInstance, DWORD dwReason, LPVOID lpReserved )
{
	if (dwReason==DLL_PROCESS_ATTACH)
	{
		wchar_t path[_MAX_PATH];
		GetModuleFileName(NULL,path,_countof(path));
		if (_wcsicmp(PathFindFileName(path),L"iexplore.exe")==0) 
			return FALSE;

		g_Instance=hInstance;

		wchar_t fname[_MAX_PATH];
		GetModuleFileName(hInstance,fname,_countof(fname));
		*PathFindFileName(fname)=0;
		wcscat_s(fname,_countof(fname),L"ExplorerL10N.ini");
		ParseSettings(fname);
		InitClassicCopyProcess();
	}

	if (dwReason==DLL_THREAD_ATTACH)
	{
		InitClassicCopyThread();
	}

	if (dwReason==DLL_THREAD_DETACH)
	{
		FreeClassicCopyThread();
	}

	return _AtlModule.DllMain(dwReason, lpReserved); 
}
