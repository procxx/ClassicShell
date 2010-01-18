// Classic Shell (c) 2009-2010, Ivo Beltchev
// The sources for Classic Shell are distributed under the MIT open source license

// dllmain.cpp : Defines the entry point for the DLL application.

#include "stdafx.h"
#include "GlobalSettings.h"
#include "TranslationSettings.h"
#include "ClassicStartMenuDLL.h"
#include "IconManager.h"
#include "Settings.h"

#pragma comment(linker, \
	"\"/manifestdependency:type='Win32' "\
	"name='Microsoft.Windows.Common-Controls' "\
	"version='6.0.0.0' "\
	"processorArchitecture='*' "\
	"publicKeyToken='6595b64144ccf1df' "\
	"language='*'\"")

BOOL APIENTRY DllMain( HMODULE hModule, DWORD ul_reason_for_call, LPVOID lpReserved )
{
	switch (ul_reason_for_call)
	{
	case DLL_PROCESS_ATTACH:
		// enable these to track heap problems
//		_CrtSetDbgFlag( _CRTDBG_ALLOC_MEM_DF | _CRTDBG_LEAK_CHECK_DF | _CRTDBG_CHECK_ALWAYS_DF );
//		_CrtSetBreakAlloc(91);
		g_Instance=hModule;
		{
#ifdef BUILD_SETUP
#define INI_PATH L""
#else
#define INI_PATH L"..\\"
#endif

			wchar_t fname[_MAX_PATH];
			GetModuleFileName(hModule,fname,_countof(fname));
			*PathFindFileName(fname)=0;
			wcscat_s(fname,_countof(fname),INI_PATH L"StartMenu.ini");
			ParseGlobalSettings(fname);

			GetModuleFileName(hModule,fname,_countof(fname));
			*PathFindFileName(fname)=0;
			wcscat_s(fname,_countof(fname),INI_PATH L"StartMenuL10N.ini");
			ParseTranslations(fname);
			g_IconManager.Init();
		}
		break;
	case DLL_PROCESS_DETACH:
		if (g_OwnerWindow) DestroyWindow(g_OwnerWindow);
		CloseSettings();
		// just in case the drop target is not already unhooked, which shouldn't happen
		// this is not really safe doing here because we don't know which thread this will execute in
		// but it's better than nothing
		UnhookDropTarget();
		// same for stopping the preloading thread. not really safe but necessary when the DLL is unloaded abnormally
		g_IconManager.StopPreloading(false);
		break;
	case DLL_THREAD_ATTACH:
	case DLL_THREAD_DETACH:
		break;
	}
	return TRUE;
}
