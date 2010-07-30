// Classic Shell (c) 2009-2010, Ivo Beltchev
// The sources for Classic Shell are distributed under the MIT open source license

// dllmain.cpp : Defines the entry point for the DLL application.

#include "stdafx.h"
#include "GlobalSettings.h"
#include "TranslationSettings.h"
#include "ClassicStartMenuDLL.h"
#include "IconManager.h"
#include "Settings.h"
#include "StringSet.h"
#include "resource.h"

#pragma comment(linker, \
	"\"/manifestdependency:type='Win32' "\
	"name='Microsoft.Windows.Common-Controls' "\
	"version='6.0.0.0' "\
	"processorArchitecture='*' "\
	"publicKeyToken='6595b64144ccf1df' "\
	"language='*'\"")

static FILETIME g_IniTimestamp;
static CRITICAL_SECTION g_IniSection;
static CStringSet g_ResStrings;
static std::vector<char> g_DlgSettings;

void ReadIniFile( void )
{
	wchar_t fname[_MAX_PATH];
	GetModuleFileName(g_Instance,fname,_countof(fname));
	*PathFindFileName(fname)=0;
	Strcat(fname,_countof(fname),INI_PATH L"StartMenu.ini");
	WIN32_FILE_ATTRIBUTE_DATA data;
	if (GetFileAttributesEx(fname,GetFileExInfoStandard,&data))
	{
		if (CompareFileTime(&g_IniTimestamp,&data.ftLastWriteTime)!=0)
		{
			g_IniTimestamp=data.ftLastWriteTime;
			ParseGlobalSettings(fname);
		}
	}
}

CString LoadStringEx( int stringID )
{
	CString str=g_ResStrings.GetString(stringID);
	if (str.IsEmpty())
		str.LoadString(g_Instance,stringID);
	return str;
}

HWND CreateSettingsDialog( HWND hWndParent, DLGPROC lpDialogFunc )
{
	if (!g_DlgSettings.empty())
	{
		HWND res=CreateDialogIndirect(g_Instance,(DLGTEMPLATE*)&g_DlgSettings[0],hWndParent,lpDialogFunc);
		if (res) return res;
	}
	return CreateDialog(g_Instance,MAKEINTRESOURCE(IDD_SETTINGS),hWndParent,lpDialogFunc);
}

DWORD GetVersionEx( HINSTANCE hInstance )
{
	// get the DLL version. this is a bit hacky. the standard way is to use GetFileVersionInfo and such API.
	// but it takes a file name instead of module handle so it will probably load the DLL a second time.
	// the header of the version resource is a fixed size so we can count on VS_FIXEDFILEINFO to always
	// be at offset 40
	HRSRC hResInfo=FindResource(hInstance,MAKEINTRESOURCE(VS_VERSION_INFO),RT_VERSION);
	if (!hResInfo)
		return 0;
	HGLOBAL hRes=LoadResource(hInstance,hResInfo);
	void *pRes=LockResource(hRes);
	if (!pRes) return 0;

	VS_FIXEDFILEINFO *pVer=(VS_FIXEDFILEINFO*)((char*)pRes+40);
	return ((HIWORD(pVer->dwProductVersionMS)&255)<<24)|((LOWORD(pVer->dwProductVersionMS)&255)<<16)|HIWORD(pVer->dwProductVersionLS);
}

extern "C" BOOL WINAPI DllMain( HINSTANCE hInstance, DWORD dwReason, LPVOID lpReserved )
{
	if (dwReason==DLL_PROCESS_ATTACH)
	{
		g_Instance=hInstance;
		ReadIniFile();

		wchar_t path[_MAX_PATH];
		GetModuleFileName(hInstance,path,_countof(path));
		*PathFindFileName(path)=0;

		wchar_t fname[_MAX_PATH];
		Sprintf(fname,_countof(fname),L"%s" INI_PATH L"StartMenuL10N.ini",path);
		const wchar_t *language=FindSetting("Language");
		ParseTranslations(fname,language);

		HINSTANCE resInstance=NULL;
		if (language)
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

			for (language=languages;*language;language+=Strlen(language)+1)
			{
				wchar_t fname[_MAX_PATH];
				Sprintf(fname,_countof(fname),L"%s" INI_PATH L"%s.dll",path,language);
				resInstance=LoadLibraryEx(fname,NULL,LOAD_LIBRARY_AS_DATAFILE|LOAD_LIBRARY_AS_IMAGE_RESOURCE);
				if (resInstance)
					break;
			}
		}

		if (resInstance)
		{
			if (GetVersionEx(resInstance)==GetVersionEx(g_Instance))
			{
				g_ResStrings.Init(resInstance);

				// load the IDD_SETTINGS dialog
				HRSRC hrsrc=FindResource(resInstance,MAKEINTRESOURCE(IDD_SETTINGS),RT_DIALOG);
				if (hrsrc)
				{
					HGLOBAL hglb=LoadResource(resInstance,hrsrc);
					if (hglb)
					{
						// finally lock the resource
						LPVOID res=LockResource(hglb);
						g_DlgSettings.resize(SizeofResource(resInstance,hrsrc));
						if (!g_DlgSettings.empty())
							memcpy(&g_DlgSettings[0],res,g_DlgSettings.size());
					}
				}

			}
			FreeLibrary(resInstance);
		}

		g_IconManager.Init();
	}

	if (dwReason==DLL_PROCESS_DETACH)
	{
		g_ResStrings.clear();
		g_DlgSettings.clear();
	}

	return TRUE;
}
