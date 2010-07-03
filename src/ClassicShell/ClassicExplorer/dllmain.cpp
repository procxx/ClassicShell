// Classic Shell (c) 2009-2010, Ivo Beltchev
// The sources for Classic Shell are distributed under the MIT open source license

// dllmain.cpp : Implementation of DllMain.

#include "stdafx.h"
#include "resource.h"
#include "dllmain.h"
#include "GlobalSettings.h"
#include "TranslationSettings.h"
#include "ShareOverlay.h"

CClassicExplorerModule _AtlModule;

void InitClassicCopyProcess( void );
void InitClassicCopyThread( void );
void FreeClassicCopyThread( void );

bool g_bHookCopyThreads;
bool g_bExplorerExe;
static FILETIME g_IniTimestamp;
static CRITICAL_SECTION g_IniSection;

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

HICON CreateDisabledIcon( HICON icon, int size )
{
	// convert normal icon to grayscale
	ICONINFO info;
	GetIconInfo(icon,&info);

	BITMAPINFO bi={0};
	bi.bmiHeader.biSize=sizeof(BITMAPINFOHEADER);
	bi.bmiHeader.biWidth=bi.bmiHeader.biHeight=size;
	bi.bmiHeader.biPlanes=1;
	bi.bmiHeader.biBitCount=32;
	HDC hdc=CreateCompatibleDC(NULL);

	unsigned int *bits;
	HBITMAP bmp=CreateDIBSection(hdc,&bi,DIB_RGB_COLORS,(void**)&bits,NULL,0);
	HGDIOBJ bmp0=SelectObject(hdc,bmp);
	RECT rc={0,0,size,size};
	FillRect(hdc,&rc,(HBRUSH)GetStockObject(BLACK_BRUSH));
	DrawIconEx(hdc,0,0,icon,size,size,0,NULL,DI_NORMAL);
	SelectObject(hdc,bmp0);
	DeleteDC(hdc);
	if (info.hbmColor) DeleteObject(info.hbmColor);
	info.hbmColor=bmp;
	int n=size*size;
	for (int i=0;i<n;i++)
	{
		unsigned int &pixel=bits[i];
		int r=(pixel&255);
		int g=((pixel>>8)&255);
		int b=((pixel>>16)&255);
		int l=(77*r+151*g+28*b)/256;
		pixel=(pixel&0xFF000000)|(l*0x010101);
	}

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

// DLL Entry Point
extern "C" BOOL WINAPI DllMain( HINSTANCE hInstance, DWORD dwReason, LPVOID lpReserved )
{
	if (dwReason==DLL_PROCESS_ATTACH)
	{
		DWORD EnableCopyUI=1;
		DWORD SharedOverlay=0;
		CRegKey regSettings;
		if (regSettings.Open(HKEY_CURRENT_USER,L"Software\\IvoSoft\\ClassicExplorer")==ERROR_SUCCESS)
		{
			regSettings.QueryDWORDValue(L"EnableCopyUI",EnableCopyUI);
			regSettings.QueryDWORDValue(L"SharedOverlay",SharedOverlay);
		}

		g_TlsIndex=TlsAlloc();
		if (g_TlsIndex==TLS_OUT_OF_INDEXES) 
			return FALSE; // TLS failure

		wchar_t path[_MAX_PATH];
		GetModuleFileName(NULL,path,_countof(path));
		const wchar_t *exe=PathFindFileName(path);
		g_bExplorerExe=(_wcsicmp(exe,L"explorer.exe")==0 || _wcsicmp(exe,L"verclsid.exe")==0);
		if (_wcsicmp(exe,L"regsvr32.exe")!=0 && _wcsicmp(exe,L"msiexec.exe")!=0 && !g_bExplorerExe && SharedOverlay!=2 && !(EnableCopyUI&4))
			return FALSE;

		g_Instance=hInstance;

		InitializeCriticalSection(&g_IniSection);
		ReadIniFile(true);

		wchar_t fname[_MAX_PATH];
		GetModuleFileName(hInstance,fname,_countof(fname));
		*PathFindFileName(fname)=0;
		Strcat(fname,_countof(fname),INI_PATH L"ExplorerL10N.ini");
		ParseTranslations(fname);

		g_bHookCopyThreads=(((EnableCopyUI&4) || g_bExplorerExe) && ((EnableCopyUI&3)==1 || (EnableCopyUI&3)==2));
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
	}

	return _AtlModule.DllMain(dwReason, lpReserved);
}
