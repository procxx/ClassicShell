// Classic Shell (c) 2009-2012, Ivo Beltchev
// The sources for Classic Shell are distributed under the MIT open source license

#include "stdafx.h"
#include "resource.h"
#include "..\..\ClassicShellLib\resource.h"
#include "Settings.h"
#include "SettingsUI.h"
#include "Translations.h"
#include "ResourceHelper.h"
#include "dllmain.h"
#include "ClassicIE9DLL.h"

#pragma comment(linker, \
	"\"/manifestdependency:type='Win32' "\
	"name='Microsoft.Windows.Common-Controls' "\
	"version='6.0.0.0' "\
	"processorArchitecture='*' "\
	"publicKeyToken='6595b64144ccf1df' "\
	"language='*'\"")

CClassicIE9DLLModule _AtlModule;

static int g_LoadDialogs[]=
{
	IDD_SETTINGS,
	IDD_SETTINGSTREE,
	IDD_LANGUAGE,
	0
};

const wchar_t *GetDocRelativePath( void )
{
	return DOC_PATH;
}

// DLL Entry Point
extern "C" BOOL WINAPI DllMain( HINSTANCE hInstance, DWORD dwReason, LPVOID lpReserved )
{
	if (dwReason==DLL_PROCESS_ATTACH)
	{
		wchar_t path[_MAX_PATH];
		GetModuleFileName(NULL,path,_countof(path));
		const wchar_t *exe=PathFindFileName(path);
		if (_wcsicmp(exe,L"explorer.exe")==0) return FALSE;
		bool bIE9=false;
		if (_wcsicmp(exe,L"iexplore.exe")==0)
		{
			DWORD version=GetVersionEx(GetModuleHandle(NULL));
			if (version<0x09000000) return FALSE;
			bIE9=true;
		}

		g_Instance=hInstance;
		InitSettings();
		if (bIE9 && !GetSettingBool(L"ShowCaption") && !GetSettingBool(L"ShowProgress") && !GetSettingBool(L"ShowZone")) return FALSE;

		CString language=GetSettingString(L"Language");
		ParseTranslations(NULL,language);

		GetModuleFileName(hInstance,path,_countof(path));
		*PathFindFileName(path)=0;
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

			for (const wchar_t *language=languages;*language;language+=Strlen(language)+1)
			{
				wchar_t fname[_MAX_PATH];
				Sprintf(fname,_countof(fname),L"%s" INI_PATH L"%s.dll",path,language);
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
		InitClassicIE9(hInstance);
	}

	return _AtlModule.DllMain(dwReason, lpReserved); 
}
