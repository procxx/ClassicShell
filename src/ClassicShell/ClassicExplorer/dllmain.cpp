// Classic Shell (c) 2009-2010, Ivo Beltchev
// The sources for Classic Shell are distributed under the MIT open source license

// dllmain.cpp : Implementation of DllMain.

#include "stdafx.h"
#include "resource.h"
#include "ClassicExplorer_i.h"
#include "dllmain.h"
#include "GlobalSettings.h"
#include "TranslationSettings.h"

CClassicExplorerModule _AtlModule;

void InitClassicCopyProcess( void );
void InitClassicCopyThread( void );
void FreeClassicCopyThread( void );

bool g_bHookCopyThreads;

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

#ifdef BUILD_SETUP
#define INI_PATH L""
#else
#define INI_PATH L"..\\"
#endif

		wchar_t fname[_MAX_PATH];
		GetModuleFileName(hInstance,fname,_countof(fname));
		*PathFindFileName(fname)=0;
		wcscat_s(fname,_countof(fname),INI_PATH L"Explorer.ini");
		ParseGlobalSettings(fname);

		GetModuleFileName(hInstance,fname,_countof(fname));
		*PathFindFileName(fname)=0;
		wcscat_s(fname,_countof(fname),INI_PATH L"ExplorerL10N.ini");
		ParseTranslations(fname);

		DWORD EnableCopyUI=1;
		CRegKey regSettings;
		if (regSettings.Open(HKEY_CURRENT_USER,L"Software\\IvoSoft\\ClassicExplorer")==ERROR_SUCCESS)
			regSettings.QueryDWORDValue(L"EnableCopyUI",EnableCopyUI);
		g_bHookCopyThreads=(EnableCopyUI==1 || EnableCopyUI==2);

		if (g_bHookCopyThreads)
			InitClassicCopyProcess();
	}

	if (dwReason==DLL_THREAD_ATTACH)
	{
		if (g_bHookCopyThreads)
			InitClassicCopyThread();
	}

	if (dwReason==DLL_THREAD_DETACH)
	{
		if (g_bHookCopyThreads)
			FreeClassicCopyThread();
	}

	return _AtlModule.DllMain(dwReason, lpReserved); 
}
