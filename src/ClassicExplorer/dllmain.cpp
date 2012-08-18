// Classic Shell (c) 2009-2012, Ivo Beltchev
// The sources for Classic Shell are distributed under the MIT open source license

// dllmain.cpp : Implementation of DllMain.

#include "stdafx.h"
#include "resource.h"
#include "..\ClassicShellLib\resource.h"
#include "dllmain.h"
#include "ShareOverlay.h"
#include "SettingsUI.h"
#include "Settings.h"
#include "Translations.h"
#include "ResourceHelper.h"
#include "FNVHash.h"
#include <uxtheme.h>
#include <dwmapi.h>

CClassicExplorerModule _AtlModule;

void InitClassicCopyProcess( void );
void InitClassicCopyThread( void );
void FreeClassicCopyThread( void );

bool g_bHookCopyThreads;
bool g_bExplorerExe;
LPCWSTR g_LoadedSettingsAtom;

static int g_LoadDialogs[]=
{
	IDD_SETTINGS,
	IDD_SETTINGSTREE,
	IDD_BROWSEFORICON,
	IDD_LANGUAGE,
	IDD_CUSTOMTOOLBAR,
	IDD_CUSTOMTREE,
	0
};

const wchar_t *GetDocRelativePath( void )
{
	return DOC_PATH;
}

struct FindChild
{
	const wchar_t *className;
	HWND hWnd;
};

static BOOL CALLBACK EnumChildProc( HWND hwnd, LPARAM lParam )
{
	FindChild &find=*(FindChild*)lParam;
	wchar_t name[256];
	GetClassName(hwnd,name,_countof(name));
	if (_wcsicmp(name,find.className)!=0) return TRUE;
	find.hWnd=hwnd;
	return FALSE;
}

HWND FindChildWindow( HWND hwnd, const wchar_t *className )
{
	FindChild find={className};
	EnumChildWindows(hwnd,EnumChildProc,(LPARAM)&find);
	return find.hWnd;
}

static DWORD g_TlsIndex;

TlsData *GetTlsData( void )
{
	void *pData=TlsGetValue(g_TlsIndex);
	if (!pData)
	{
		pData=(void*)LocalAlloc(LPTR,sizeof(TlsData));
		memset(pData,0,sizeof(TlsData));
		TlsSetValue(g_TlsIndex,pData);
	}
	return (TlsData*)pData;
}

// DLL Entry Point
extern "C" BOOL WINAPI DllMain( HINSTANCE hInstance, DWORD dwReason, LPVOID lpReserved )
{
	if (dwReason==DLL_PROCESS_ATTACH)
	{
		InitSettings();
		g_TlsIndex=TlsAlloc();
		if (g_TlsIndex==TLS_OUT_OF_INDEXES) 
			return FALSE; // TLS failure

		wchar_t path[_MAX_PATH];
		GetModuleFileName(NULL,path,_countof(path));
		const wchar_t *exe=PathFindFileName(path);
		g_bExplorerExe=(_wcsicmp(exe,L"explorer.exe")==0 || _wcsicmp(exe,L"verclsid.exe")==0);
		bool bReplaceUI=GetWinVersion()<WIN_VER_WIN8 && (GetSettingBool(L"ReplaceFileUI") || GetSettingBool(L"ReplaceFolderUI") || GetSettingBool(L"EnableMore"));
		if (_wcsicmp(exe,L"regsvr32.exe")!=0 && _wcsicmp(exe,L"msiexec.exe")!=0 && _wcsicmp(exe,L"ClassicExplorerSettings.exe")!=0 && !g_bExplorerExe)
		{
			// some arbitrary app
			if ((!GetSettingBool(L"ShareOverlay") || GetSettingBool(L"ShareExplorer")) && (!bReplaceUI || GetSettingBool(L"FileExplorer")))
				return FALSE;
			CString whiteList=GetSettingString(L"ProcessWhiteList");
			if (!whiteList.IsEmpty())
			{
				// check for whitelisted process names
				const wchar_t *str=whiteList;
				bool bFound=false;
				while (*str)
				{
					wchar_t token[_MAX_PATH];
					str=GetToken(str,token,_countof(token),L",;");
					wchar_t *start=token;
					while (*start==' ')
						start++;
					wchar_t *end=start+Strlen(start);
					while (end>start && end[-1]==' ')
						end--;
					*end=0;
					if (_wcsicmp(exe,start)==0)
					{
						bFound=true;
						break;
					}
				}
				if (!bFound)
					return FALSE;
			}
			else
			{
				// check for blacklisted process names
				CString blackList=GetSettingString(L"ProcessBlackList");
				const wchar_t *str=blackList;
				while (*str)
				{
					wchar_t token[_MAX_PATH];
					str=GetToken(str,token,_countof(token),L",;");
					wchar_t *start=token;
					while (*start==' ')
						start++;
					wchar_t *end=start+Strlen(start);
					while (end>start && end[-1]==' ')
						end--;
					*end=0;
					if (_wcsicmp(exe,start)==0)
						return FALSE;
				}
			}
		}

		g_Instance=hInstance;
		g_LoadedSettingsAtom=(LPCWSTR)GlobalAddAtom(L"ClassicExplorer.LoadedSettings");

		GetModuleFileName(hInstance,path,_countof(path));
		*PathFindFileName(path)=0;
		wchar_t fname[_MAX_PATH];
		Sprintf(fname,_countof(fname),L"%s" INI_PATH L"ExplorerL10N.ini",path);
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

		g_bHookCopyThreads=(bReplaceUI && (g_bExplorerExe || !GetSettingBool(L"FileExplorer")));
		if (g_bHookCopyThreads)
		{
			InitClassicCopyProcess();
			InitClassicCopyThread();
		}

		if (GetSettingBool(L"ShareOverlay") && (g_bExplorerExe || !GetSettingBool(L"ShareExplorer")))
			CShareOverlay::InitOverlay(GetSettingString(L"ShareOverlayIcon"));
	}

	if (dwReason==DLL_THREAD_ATTACH)
	{
		if (g_bHookCopyThreads)
			InitClassicCopyThread();
	}

	if (dwReason==DLL_THREAD_DETACH)
	{
		void *pData=TlsGetValue(g_TlsIndex);
		if (pData)
			LocalFree((HLOCAL)pData);
		TlsSetValue(g_TlsIndex,NULL);
		if (g_bHookCopyThreads)
			FreeClassicCopyThread();
	}

	if (dwReason==DLL_PROCESS_DETACH)
	{
		void *pData=TlsGetValue(g_TlsIndex);
		if (pData)
			LocalFree((HLOCAL)pData);
		TlsSetValue(g_TlsIndex,NULL);
		TlsFree(g_TlsIndex);
		GlobalDeleteAtom((ATOM)g_LoadedSettingsAtom);
	}

	return _AtlModule.DllMain(dwReason, lpReserved);
}
