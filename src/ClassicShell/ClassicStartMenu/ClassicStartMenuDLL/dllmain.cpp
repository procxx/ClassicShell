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
	if (ul_reason_for_call==DLL_PROCESS_ATTACH)
	{
		g_Instance=hModule;
		wchar_t fname[_MAX_PATH];
		GetModuleFileName(hModule,fname,_countof(fname));
		*PathFindFileName(fname)=0;
		Strcat(fname,_countof(fname),INI_PATH L"StartMenu.ini");
		ParseGlobalSettings(fname);

		GetModuleFileName(hModule,fname,_countof(fname));
		*PathFindFileName(fname)=0;
		Strcat(fname,_countof(fname),INI_PATH L"StartMenuL10N.ini");
		ParseTranslations(fname);
		g_IconManager.Init();
	}

	return TRUE;
}
