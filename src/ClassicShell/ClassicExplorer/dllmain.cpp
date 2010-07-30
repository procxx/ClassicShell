// Classic Shell (c) 2009-2010, Ivo Beltchev
// The sources for Classic Shell are distributed under the MIT open source license

// dllmain.cpp : Implementation of DllMain.

#include "stdafx.h"
#include "resource.h"
#include "dllmain.h"
#include "GlobalSettings.h"
#include "TranslationSettings.h"
#include "ShareOverlay.h"
#include "StringSet.h"

CClassicExplorerModule _AtlModule;

void InitClassicCopyProcess( void );
void InitClassicCopyThread( void );
void FreeClassicCopyThread( void );

bool g_bHookCopyThreads;
bool g_bExplorerExe;
static FILETIME g_IniTimestamp;
static CRITICAL_SECTION g_IniSection;

static CStringSet g_ResStrings;
static std::vector<char> g_DlgSettings;

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

void ReadIniFile( bool bStartup )
{
	if (!bStartup)
		EnterCriticalSection(&g_IniSection);
	wchar_t fname[_MAX_PATH];
	GetModuleFileName(g_Instance,fname,_countof(fname));
	*PathFindFileName(fname)=0;
	Strcat(fname,_countof(fname),INI_PATH L"Explorer.ini");
	WIN32_FILE_ATTRIBUTE_DATA data;
	if (GetFileAttributesEx(fname,GetFileExInfoStandard,&data))
	{
		if (CompareFileTime(&g_IniTimestamp,&data.ftLastWriteTime)!=0)
		{
			g_IniTimestamp=data.ftLastWriteTime;
			ParseGlobalSettings(fname);
		}
	}
	if (!bStartup)
		LeaveCriticalSection(&g_IniSection);
}

HICON LoadIcon( int iconSize, const wchar_t *path, int index, std::vector<HMODULE> &modules, HMODULE hShell32 )
{
	if (!path)
		return (HICON)LoadImage(hShell32,MAKEINTRESOURCE(index),IMAGE_ICON,iconSize,iconSize,LR_DEFAULTCOLOR);
	wchar_t text[1024];
	Strcpy(text,_countof(text),path);
	DoEnvironmentSubst(text,_countof(text));
	wchar_t *c=wcsrchr(text,',');
	if (c)
	{
		// resource file
		*c=0;
		const wchar_t *res=c+1;
		int idx=_wtol(res);
		if (idx>0) res=MAKEINTRESOURCE(idx);
		if (!text[0])
			return (HICON)LoadImage(g_Instance,res,IMAGE_ICON,iconSize,iconSize,LR_DEFAULTCOLOR);
		HMODULE hMod=GetModuleHandle(PathFindFileName(text));
		if (!hMod)
		{
			hMod=LoadLibraryEx(text,NULL,LOAD_LIBRARY_AS_DATAFILE|LOAD_LIBRARY_AS_IMAGE_RESOURCE);
			if (!hMod) return NULL;
			modules.push_back(hMod);
		}
		return (HICON)LoadImage(hMod,res,IMAGE_ICON,iconSize,iconSize,LR_DEFAULTCOLOR);
	}
	else
	{
		return (HICON)LoadImage(NULL,text,IMAGE_ICON,iconSize,iconSize,LR_DEFAULTCOLOR|LR_LOADFROMFILE);
	}
}

HBITMAP BitmapFromIcon( HICON icon, int iconSize, unsigned int **pBits, bool bDestroyIcon )
{
	BITMAPINFO bi={0};
	bi.bmiHeader.biSize=sizeof(BITMAPINFOHEADER);
	bi.bmiHeader.biWidth=bi.bmiHeader.biHeight=iconSize;
	bi.bmiHeader.biPlanes=1;
	bi.bmiHeader.biBitCount=32;
	RECT rc={0,0,iconSize,iconSize};

	HDC hdc=CreateCompatibleDC(NULL);
	unsigned int *bits;
	HBITMAP bmp=CreateDIBSection(hdc,&bi,DIB_RGB_COLORS,(void**)&bits,NULL,0);
	HGDIOBJ bmp0=SelectObject(hdc,bmp);
	FillRect(hdc,&rc,(HBRUSH)GetStockObject(BLACK_BRUSH));
	DrawIconEx(hdc,0,0,icon,iconSize,iconSize,0,NULL,DI_NORMAL);
	SelectObject(hdc,bmp0);
	if (bDestroyIcon) DestroyIcon(icon);
	if (pBits) *pBits=bits;
	return bmp;
}

HICON CreateDisabledIcon( HICON icon, int iconSize )
{
	// convert normal icon to grayscale
	ICONINFO info;
	GetIconInfo(icon,&info);

	unsigned int *bits;
	HBITMAP bmp=BitmapFromIcon(icon,iconSize,&bits,false);

	int n=iconSize*iconSize;
	for (int i=0;i<n;i++)
	{
		unsigned int &pixel=bits[i];
		int r=(pixel&255);
		int g=((pixel>>8)&255);
		int b=((pixel>>16)&255);
		int l=(77*r+151*g+28*b)/256;
		pixel=(pixel&0xFF000000)|(l*0x010101);
	}

	if (info.hbmColor) DeleteObject(info.hbmColor);
	info.hbmColor=bmp;
	icon=CreateIconIndirect(&info);
	DeleteObject(bmp);
	if (info.hbmMask) DeleteObject(info.hbmMask);
	return icon;
}

static DWORD g_TlsIndex;

TlsData *GetTlsData( void )
{
	void *pData=TlsGetValue(g_TlsIndex);
	if (!pData)
	{
		pData=(void*)LocalAlloc(LPTR,sizeof(TlsData));
		TlsSetValue(g_TlsIndex,pData);
	}
	return (TlsData*)pData;
}

CString LoadStringEx( int stringID )
{
	CString str=g_ResStrings.GetString(stringID);
	if (str.IsEmpty())
		str.LoadString(g_Instance,stringID);
	return str;
}

INT_PTR RunSettingsDialog( HWND hWndParent, DLGPROC lpDialogFunc )
{
	if (g_DlgSettings.empty())
		return DialogBox(g_Instance,MAKEINTRESOURCE(IDD_SETTINGS),hWndParent,lpDialogFunc);
	else
		return DialogBoxIndirect(g_Instance,(DLGTEMPLATE*)&g_DlgSettings[0],hWndParent,lpDialogFunc);
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

// DLL Entry Point
extern "C" BOOL WINAPI DllMain( HINSTANCE hInstance, DWORD dwReason, LPVOID lpReserved )
{
	if (dwReason==DLL_PROCESS_ATTACH)
	{
		DWORD EnableFileUI=FILEUI_DEFAULT;
		DWORD SharedOverlay=0;
		CRegKey regSettings;
		if (regSettings.Open(HKEY_CURRENT_USER,L"Software\\IvoSoft\\ClassicExplorer")==ERROR_SUCCESS)
		{
			regSettings.QueryDWORDValue(L"EnableFileUI",EnableFileUI);
			regSettings.QueryDWORDValue(L"SharedOverlay",SharedOverlay);
		}

		g_TlsIndex=TlsAlloc();
		if (g_TlsIndex==TLS_OUT_OF_INDEXES) 
			return FALSE; // TLS failure

		wchar_t path[_MAX_PATH];
		GetModuleFileName(NULL,path,_countof(path));
		const wchar_t *exe=PathFindFileName(path);
		g_bExplorerExe=(_wcsicmp(exe,L"explorer.exe")==0 || _wcsicmp(exe,L"verclsid.exe")==0);
		if (_wcsicmp(exe,L"regsvr32.exe")!=0 && _wcsicmp(exe,L"msiexec.exe")!=0 && _wcsicmp(exe,L"ClassicExplorerSettings.exe")!=0 && !g_bExplorerExe && SharedOverlay!=2 && !((EnableFileUI&FILEUI_OTHERAPPS) && (EnableFileUI&FILEUI_ALL)))
			return FALSE;

		g_Instance=hInstance;

		InitializeCriticalSection(&g_IniSection);
		ReadIniFile(true);

		GetModuleFileName(hInstance,path,_countof(path));
		*PathFindFileName(path)=0;
		wchar_t fname[_MAX_PATH];
		Sprintf(fname,_countof(fname),L"%s" INI_PATH L"ExplorerL10N.ini",path);
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

		g_bHookCopyThreads=(((EnableFileUI&FILEUI_OTHERAPPS) || g_bExplorerExe) && (EnableFileUI&FILEUI_ALL));
		if (g_bHookCopyThreads)
		{
			InitClassicCopyProcess();
			InitClassicCopyThread();
		}

		if ((g_bExplorerExe && SharedOverlay) || SharedOverlay==2)
			CShareOverlay::InitOverlay(FindSetting("ShareOverlayIcon"));
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
		DeleteCriticalSection(&g_IniSection);
		g_ResStrings.clear();
		g_DlgSettings.clear();
	}

	return _AtlModule.DllMain(dwReason, lpReserved);
}
