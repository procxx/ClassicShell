// Classic Shell (c) 2009-2010, Ivo Beltchev
// The sources for Classic Shell are distributed under the MIT open source license

// dllmain.cpp : Defines the entry point for the DLL application.

#include "stdafx.h"
#include "ClassicStartMenuDLL.h"
#include "IconManager.h"
#include "Settings.h"
#include "Translations.h"
#include "ResourceHelper.h"
#include "StringSet.h"
#include "resource.h"
#include "..\ClassicShellLib\resource.h"
#include "SettingsUI.h"
#include "SkinManager.h"
#include "uxtheme.h"
#include "FNVHash.h"
#include "MenuContainer.h"
#include <dwmapi.h>

#pragma comment(linker, \
	"\"/manifestdependency:type='Win32' "\
	"name='Microsoft.Windows.Common-Controls' "\
	"version='6.0.0.0' "\
	"processorArchitecture='*' "\
	"publicKeyToken='6595b64144ccf1df' "\
	"language='*'\"")

static int g_LoadDialogs[]=
{
	IDD_SETTINGS,
	IDD_SETTINGSTREE,
	IDD_BROWSEFORICON,
	IDD_LANGUAGE,
	IDD_SKINSETTINGS,
	IDD_CUSTOMTREE,
	IDD_CUSTOMMENU,
	0
};

const wchar_t *GetDocRelativePath( void )
{
	return DOC_PATH;
}

extern "C" BOOL WINAPI DllMain( HINSTANCE hInstance, DWORD dwReason, LPVOID lpReserved )
{
	if (dwReason==DLL_PROCESS_ATTACH)
	{
		g_Instance=hInstance;
		InitSettings();

		wchar_t path[_MAX_PATH];
		GetModuleFileName(hInstance,path,_countof(path));
		*PathFindFileName(path)=0;

		wchar_t fname[_MAX_PATH];
		Sprintf(fname,_countof(fname),L"%s" INI_PATH L"StartMenuL10N.ini",path);
		CString language=GetSettingString(L"Language");
		ParseTranslations(fname,language);

		HINSTANCE resInstance=NULL;
		if (!language.IsEmpty())
		{
			wchar_t fname[_MAX_PATH];
			Sprintf(fname,_countof(fname),L"%s" INI_PATH L"%s.dll",path,language);
			resInstance=LoadLibraryEx(fname,NULL,LOAD_LIBRARY_AS_DATAFILE|LOAD_LIBRARY_AS_IMAGE_RESOURCE);
		}
		else
		{
			wchar_t languages[100]={0};
			ULONG size=4; // up to 4 languages
			ULONG len=_countof(languages);
			GetThreadPreferredUILanguages(MUI_LANGUAGE_NAME,&size,languages,&len);

			for (const wchar_t *lang=languages;*lang;lang+=Strlen(lang)+1)
			{
				wchar_t fname[_MAX_PATH];
				Sprintf(fname,_countof(fname),L"%s" INI_PATH L"%s.dll",path,lang);
				resInstance=LoadLibraryEx(fname,NULL,LOAD_LIBRARY_AS_DATAFILE|LOAD_LIBRARY_AS_IMAGE_RESOURCE);
				if (resInstance)
					break;
			}
		}

		if (resInstance && GetVersionEx(resInstance)!=GetVersionEx(g_Instance))
		{
			FreeLibrary(resInstance);
			resInstance=NULL;
		}
		LoadTranslationResources(g_Instance,resInstance,g_LoadDialogs);

		if (resInstance)
			FreeLibrary(resInstance);

		g_IconManager.Init();
	}

	return TRUE;
}
