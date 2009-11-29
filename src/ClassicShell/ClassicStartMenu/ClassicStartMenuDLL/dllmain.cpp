// Classic Shell (c) 2009, Ivo Beltchev
// The sources for Classic Shell are distributed under the MIT open source license

// dllmain.cpp : Defines the entry point for the DLL application.

#include "stdafx.h"
#include "..\LocalizationSettings\ParseSettings.h"
#include "ClassicStartMenuDLL.h"
#include "IconManager.h"
#include "Settings.h"

BOOL APIENTRY DllMain( HMODULE hModule,
                       DWORD  ul_reason_for_call,
                       LPVOID lpReserved
					 )
{
	switch (ul_reason_for_call)
	{
	case DLL_PROCESS_ATTACH:
		// enable these to track heap problems
//		_CrtSetDbgFlag( _CRTDBG_ALLOC_MEM_DF | _CRTDBG_LEAK_CHECK_DF | _CRTDBG_CHECK_ALWAYS_DF );
//		_CrtSetBreakAlloc(91);
		g_Instance=hModule;
		{
			wchar_t fname[_MAX_PATH];
			GetModuleFileName(hModule,fname,_countof(fname));
			*PathFindFileName(fname)=0;
			wcscat_s(fname,_countof(fname),L"StartMenuL10N.ini");
			ParseSettings(fname);
		}
		break;
	case DLL_PROCESS_DETACH:
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

