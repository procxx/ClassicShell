// Classic Shell (c) 2009-2010, Ivo Beltchev
// The sources for Classic Shell are distributed under the MIT open source license

#include "stdafx.h"
#include "IconManager.h"
#include "FNVHash.h"
#include "GlobalSettings.h"
#include "TranslationSettings.h"

const int MAX_FOLDER_LEVEL=10; // don't go more than 10 levels deep

int CIconManager::s_DPI;
int CIconManager::LARGE_ICON_SIZE;
int CIconManager::SMALL_ICON_SIZE;
bool CIconManager::s_bStopLoading;
std::map<unsigned int,HICON> CIconManager::s_PreloadedIcons;
CRITICAL_SECTION CIconManager::s_PreloadSection;

CIconManager g_IconManager; // must be after s_PreloadedIcons and s_PreloadSection because the constructor starts a thread that uses the Preloaded stuff

void CIconManager::Init( void )
{
	wchar_t path[_MAX_PATH];
	GetModuleFileName(NULL,path,_countof(path));
	// the ClassicStartMenu.exe is not marked as DPI-aware. It can't call SetProcessDPIAware because it happens after
	// our DLL is loaded, and is too late. DPI-awareness can be set with a manifest, but that adds link-time warnings.
	// so we hack it here
	if (_wcsicmp(PathFindFileName(path),L"ClassicStartMenu.exe")==0)
		SetProcessDPIAware();

	{
		// get the DPI setting
		HDC hdc=::GetDC(NULL);
		s_DPI=GetDeviceCaps(hdc,LOGPIXELSY);
		::ReleaseDC(NULL,hdc);
	}

	int iconSize;
	const wchar_t *str=FindSetting("SmallIconSize");
	if (str)
	{
		iconSize=_wtol(str);
		if (iconSize<8) iconSize=8;
		if (iconSize>128) iconSize=128;
	}
	else
	{
		if (s_DPI>120)
			iconSize=24;
		else if (s_DPI>96)
			iconSize=20;
		else
			iconSize=16;
	}
	SMALL_ICON_SIZE=iconSize;
	str=FindSetting("LargeIconSize");
	if (str)
		LARGE_ICON_SIZE=_wtol(str);
	else
		LARGE_ICON_SIZE=iconSize*2;
	if (LARGE_ICON_SIZE<8) LARGE_ICON_SIZE=8;
	if (LARGE_ICON_SIZE>128) LARGE_ICON_SIZE=128;

	bool bRTL=IsLanguageRTL();
	// create the image lists
	m_LargeIcons=ImageList_Create(LARGE_ICON_SIZE,LARGE_ICON_SIZE,ILC_COLOR32|ILC_MASK|(bRTL?ILC_MIRROR:0),1,16);
	m_SmallIcons=ImageList_Create(SMALL_ICON_SIZE,SMALL_ICON_SIZE,ILC_COLOR32|ILC_MASK|(bRTL?ILC_MIRROR:0),1,16);

	// add default blank icons (default icon for file with no extension)
	SHFILEINFO info;
	if (SHGetFileInfo(L"file",FILE_ATTRIBUTE_NORMAL,&info,sizeof(info),SHGFI_USEFILEATTRIBUTES|SHGFI_ICON|SHGFI_LARGEICON))
	{
		ImageList_AddIcon(m_LargeIcons,info.hIcon);
		DestroyIcon(info.hIcon);
	}
	if (SHGetFileInfo(L"file",FILE_ATTRIBUTE_NORMAL,&info,sizeof(info),SHGFI_USEFILEATTRIBUTES|SHGFI_ICON|SHGFI_SMALLICON))
	{
		ImageList_AddIcon(m_SmallIcons,info.hIcon);
		DestroyIcon(info.hIcon);
	}

	InitializeCriticalSection(&s_PreloadSection);

	if (_wcsicmp(PathFindFileName(path),L"explorer.exe")==0)
	{
		// don't preload icons if running outside of the explorer
		s_bStopLoading=false;
		m_PreloadThread=CreateThread(NULL,0,PreloadThread,NULL,0,NULL);
	}
	else
		m_PreloadThread=INVALID_HANDLE_VALUE;
}

void CIconManager::StopPreloading( bool bWait )
{
	if (m_PreloadThread!=INVALID_HANDLE_VALUE)
	{
		s_bStopLoading=true;
		WaitForSingleObject(m_PreloadThread,bWait?INFINITE:0);
		CloseHandle(m_PreloadThread);
		DeleteCriticalSection(&s_PreloadSection);
		m_PreloadThread=INVALID_HANDLE_VALUE;
	}
}

CIconManager::~CIconManager( void )
{
	if (m_LargeIcons) ImageList_Destroy(m_LargeIcons);
	if (m_SmallIcons) ImageList_Destroy(m_SmallIcons);
}

// Retrieves an icon from a shell folder and child ID
int CIconManager::GetIcon( IShellFolder *pFolder, PCUITEMID_CHILD item, bool bLarge )
{
	ProcessPreloadedIcons();

	// get the IExtractIcon object
	CComPtr<IExtractIcon> pExtract;
	HRESULT hr=pFolder->GetUIObjectOf(NULL,1,&item,IID_IExtractIcon,NULL,(void**)&pExtract);
	HICON hIcon;
	unsigned int key;
	if (SUCCEEDED(hr))
	{
		// get the icon location
		wchar_t location[_MAX_PATH];
		int index=0;
		UINT flags=0;
		hr=pExtract->GetIconLocation(0,location,_countof(location),&index,&flags);
		if (hr!=S_OK)
			return 0;

		// check if this location+index is in the cache
		key=CalcFNVHash(location,CalcFNVHash(&index,4));
		if (bLarge)
		{
			if (m_LargeCache.find(key)!=m_LargeCache.end())
				return m_LargeCache[key];
		}
		else
		{
			EnterCriticalSection(&s_PreloadSection);
			int res=-1;
			if (m_SmallCache.find(key)!=m_SmallCache.end())
				res=m_SmallCache[key];
			LeaveCriticalSection(&s_PreloadSection);
			if (res>=0) return res;
		}

		// extract the icon
		hr=pExtract->Extract(location,index,bLarge?&hIcon:NULL,bLarge?NULL:&hIcon,MAKELONG(LARGE_ICON_SIZE,SMALL_ICON_SIZE));
		if (hr==E_INVALIDARG)
		{
			// HACK!!! There are 2 problems when extracting icons from the Games folder.
			// 1) IExtractIcon::Extract fails if the small and the large icons are not both specified even though the docs say that they are optional.
			// 2) With Microangelo On Display installed, even if both icons are specified, the function returns the default exe icon and no error code.
			//    This is probably caused by the Games shell extension supporting only Unicode, and Microangelo supporting only ANSI.
			// Fortunately because of the first problem we can detect when something is wrong (hr is E_INVALIDARG). In such case we give
			// the location/index data to ExtractIconEx, hoping it will return the correct icon. Seems to be working.
			// BTW, I have no idea how the shell doesn't have this problem. Probably because it uses IShellItems instead of PIDLs
			hr=S_FALSE;
		}
		if (hr==S_FALSE)
		{
			// the IExtractIcon object didn't do anything - use ExtractIconEx instead
			if (ExtractIconEx(location,index,bLarge?&hIcon:NULL,bLarge?NULL:&hIcon,1)==1)
				hr=S_OK;
		}
	}
	else
	{
		// try again using the ANSI version
		CComPtr<IExtractIconA> pExtractA;
		hr=pFolder->GetUIObjectOf(NULL,1,&item,IID_IExtractIconA,NULL,(void**)&pExtractA);
		if (FAILED(hr))
			return 0;

		// get the icon location
		char location[_MAX_PATH];
		int index=0;
		UINT flags=0;
		hr=pExtractA->GetIconLocation(0,location,_countof(location),&index,&flags);
		if (hr!=S_OK)
			return 0;

		// check if this location+index is in the cache
		key=CalcFNVHash(location,CalcFNVHash(&index,4));
		if (bLarge)
		{
			if (m_LargeCache.find(key)!=m_LargeCache.end())
				return m_LargeCache[key];
		}
		else
		{
			EnterCriticalSection(&s_PreloadSection);
			int res=-1;
			if (m_SmallCache.find(key)!=m_SmallCache.end())
				res=m_SmallCache[key];
			LeaveCriticalSection(&s_PreloadSection);
			if (res>=0) return res;
		}

		// extract the icon
		hr=pExtractA->Extract(location,index,bLarge?&hIcon:NULL,bLarge?NULL:&hIcon,MAKELONG(LARGE_ICON_SIZE,SMALL_ICON_SIZE));
		if (hr==E_INVALIDARG)
		{
			// HACK!!! See the previous "HACK!!!" comment for details
			hr=S_FALSE;
		}
		if (hr==S_FALSE)
		{
			// the IExtractIcon object didn't do anything - use ExtractIconEx instead
			if (ExtractIconExA(location,index,bLarge?&hIcon:NULL,bLarge?NULL:&hIcon,1)==1)
				hr=S_OK;
		}
	}

	// add to the image list
	int index=0;
	if (hr==S_OK)
	{
		index=ImageList_AddIcon(bLarge?m_LargeIcons:m_SmallIcons,hIcon);
		DestroyIcon(hIcon);
	}

	// add to the cache
	if (bLarge)
		m_LargeCache[key]=index;
	else
	{
		EnterCriticalSection(&s_PreloadSection);
		m_SmallCache[key]=index;
		LeaveCriticalSection(&s_PreloadSection);
	}

	return index;
}

// Retrieves an icon from a file and icon index (index>=0 - icon index, index<0 - resource ID)
int CIconManager::GetIcon( const wchar_t *location, int index, bool bLarge )
{
	ProcessPreloadedIcons();

	// check if this location+index is in the cache
	unsigned int key=CalcFNVHash(location,CalcFNVHash(&index,4));
	if (bLarge)
	{
		if (m_LargeCache.find(key)!=m_LargeCache.end())
			return m_LargeCache[key];
	}
	else
	{
		EnterCriticalSection(&s_PreloadSection);
		int res=-1;
		if (m_SmallCache.find(key)!=m_SmallCache.end())
			res=m_SmallCache[key];
		LeaveCriticalSection(&s_PreloadSection);
		if (res>=0) return res;
	}

	// extract icon
	HICON hIcon;
	HMODULE hMod=GetModuleHandle(PathFindFileName(location));
	if (hMod && index<0)
	{
		int iconSize=bLarge?LARGE_ICON_SIZE:SMALL_ICON_SIZE;
		hIcon=(HICON)LoadImage(hMod,MAKEINTRESOURCE(-index),IMAGE_ICON,iconSize,iconSize,LR_DEFAULTCOLOR);
	}
	else if (ExtractIconEx(location,index,bLarge?&hIcon:NULL,bLarge?NULL:&hIcon,1)!=1)
		hIcon=NULL;

	// add to the image list
	if (hIcon)
	{
		index=ImageList_AddIcon(bLarge?m_LargeIcons:m_SmallIcons,hIcon);
		DestroyIcon(hIcon);
	}
	else
		index=0;

	// add to the cache
	if (bLarge)
		m_LargeCache[key]=index;
	else
	{
		EnterCriticalSection(&s_PreloadSection);
		m_SmallCache[key]=index;
		LeaveCriticalSection(&s_PreloadSection);
	}

	return index;
}

// Retrieves an icon from shell32.dll by resource ID
int CIconManager::GetStdIcon( int id, bool bLarge )
{
	// check if this id is in the cache
	unsigned int key=CalcFNVHash(&id,4);
	if (bLarge)
	{
		if (m_LargeCache.find(key)!=m_LargeCache.end())
			return m_LargeCache[key];
	}
	else
	{
		EnterCriticalSection(&s_PreloadSection);
		int res=-1;
		if (m_SmallCache.find(key)!=m_SmallCache.end())
			res=m_SmallCache[key];
		LeaveCriticalSection(&s_PreloadSection);
		if (res>=0) return res;
	}

	// get from shell32.dll
	HMODULE hShell32=GetModuleHandle(L"shell32.dll");
	if (!hShell32) return 0;
	int size=bLarge?LARGE_ICON_SIZE:SMALL_ICON_SIZE;
	HICON hIcon=(HICON)LoadImage(hShell32,MAKEINTRESOURCE(id),IMAGE_ICON,size,size,LR_DEFAULTCOLOR);
	if (!hIcon)
		return 0;

	// add to the image list
	int index=ImageList_AddIcon(bLarge?m_LargeIcons:m_SmallIcons,hIcon);
	DestroyIcon(hIcon);

	// add to the cache
	if (bLarge)
		m_LargeCache[key]=index;
	else
	{
		EnterCriticalSection(&s_PreloadSection);
		m_SmallCache[key]=index;
		LeaveCriticalSection(&s_PreloadSection);
	}

	return index;
}

int CIconManager::GetCustomIcon( const wchar_t *path, bool bLarge )
{
	wchar_t text[1024];
	Strcpy(text,_countof(text),path);
	DoEnvironmentSubst(text,_countof(text));
	wchar_t *c=wcsrchr(text,',');
	int index=0;
	if (c)
	{
		*c=0;
		index=-_wtol(c+1);
	}
	return GetIcon(text,index,bLarge);
}

void CIconManager::ProcessPreloadedIcons( void )
{
	EnterCriticalSection(&s_PreloadSection);
	for (std::map<unsigned int,HICON>::const_iterator it=s_PreloadedIcons.begin();it!=s_PreloadedIcons.end();++it)
	{
		unsigned int key=it->first;
		HICON hIcon=it->second;
		if (m_SmallCache.find(key)==m_SmallCache.end())
		{
			// add to the image list and to the cache
			m_SmallCache[key]=ImageList_AddIcon(m_SmallIcons,hIcon);
		}
		DestroyIcon(hIcon);
	}
	s_PreloadedIcons.clear();
	LeaveCriticalSection(&s_PreloadSection);
}

// Recursive function to preload the icons for a folder
void CIconManager::LoadFolderIcons( IShellFolder *pFolder, int level )
{
	CComPtr<IEnumIDList> pEnum;
	if (pFolder->EnumObjects(NULL,SHCONTF_NONFOLDERS|SHCONTF_FOLDERS,&pEnum)!=S_OK) pEnum=NULL;
	if (!pEnum) return;

	PITEMID_CHILD pidl;
	while (pEnum->Next(1,&pidl,NULL)==S_OK)
	{
		CComPtr<IExtractIcon> pExtract;
		if (SUCCEEDED(pFolder->GetUIObjectOf(NULL,1,&pidl,IID_IExtractIcon,NULL,(void**)&pExtract)))
		{
			// get the icon location
			wchar_t location[_MAX_PATH];
			int index=0;
			UINT flags=0;
			if (SUCCEEDED(pExtract->GetIconLocation(0,location,_countof(location),&index,&flags)))
			{
				// check if this location+index is in the cache
				unsigned int key=CalcFNVHash(location,CalcFNVHash(&index,4));
				EnterCriticalSection(&s_PreloadSection);
				bool bLoaded=g_IconManager.m_SmallCache.find(key)!=g_IconManager.m_SmallCache.end() || s_PreloadedIcons.find(key)!=s_PreloadedIcons.end();
				LeaveCriticalSection(&s_PreloadSection);
				if (!bLoaded)
				{
					// extract the icon
					HICON hIcon;
					HRESULT hr=pExtract->Extract(location,index,NULL,&hIcon,MAKELONG(LARGE_ICON_SIZE,SMALL_ICON_SIZE));
					if (hr==E_INVALIDARG)
					{
						// HACK!!! See the previous "HACK!!!" comment for details
						hr=S_FALSE;
					}
					if (hr==S_FALSE)
					{
						// the IExtractIcon object didn't do anything - use ExtractIconEx instead
						if (ExtractIconEx(location,index,NULL,&hIcon,1)==1)
							hr=S_OK;
					}
					Sleep(10); // pause for a bit to reduce the stress on the system
					if (hr==S_OK)
					{
						EnterCriticalSection(&s_PreloadSection);
						s_PreloadedIcons[key]=hIcon;
						LeaveCriticalSection(&s_PreloadSection);
					}
				}
			}
		}
		else
		{
			// try the ANSI version
			CComPtr<IExtractIconA> pExtractA;
			if (SUCCEEDED(pFolder->GetUIObjectOf(NULL,1,&pidl,IID_IExtractIconA,NULL,(void**)&pExtractA)))
			{
				// get the icon location
				char location[_MAX_PATH];
				int index=0;
				UINT flags=0;
				if (SUCCEEDED(pExtractA->GetIconLocation(0,location,_countof(location),&index,&flags)))
				{
					// check if this location+index is in the cache
					unsigned int key=CalcFNVHash(location,CalcFNVHash(&index,4));
					EnterCriticalSection(&s_PreloadSection);
					bool bLoaded=g_IconManager.m_SmallCache.find(key)!=g_IconManager.m_SmallCache.end() || s_PreloadedIcons.find(key)!=s_PreloadedIcons.end();
					LeaveCriticalSection(&s_PreloadSection);
					if (!bLoaded)
					{
						// extract the icon
						HICON hIcon;
						HRESULT hr=pExtractA->Extract(location,index,NULL,&hIcon,MAKELONG(LARGE_ICON_SIZE,SMALL_ICON_SIZE));
						if (hr==E_INVALIDARG)
						{
							// HACK!!! See the previous "HACK!!!" comment for details
							hr=S_FALSE;
						}
						if (hr==S_FALSE)
						{
							// the IExtractIcon object didn't do anything - use ExtractIconEx instead
							if (ExtractIconExA(location,index,NULL,&hIcon,1)==1)
								hr=S_OK;
						}
						Sleep(10); // pause for a bit to reduce the stress on the system
						if (hr==S_OK)
						{
							EnterCriticalSection(&s_PreloadSection);
							s_PreloadedIcons[key]=hIcon;
							LeaveCriticalSection(&s_PreloadSection);
						}
					}
				}
			}
		}

		if (level<MAX_FOLDER_LEVEL)
		{
			SFGAOF flags=SFGAO_FOLDER|SFGAO_STREAM|SFGAO_LINK;
			if (SUCCEEDED(pFolder->GetAttributesOf(1,&pidl,&flags)) && (flags&(SFGAO_FOLDER|SFGAO_STREAM|SFGAO_LINK))==SFGAO_FOLDER)
			{
				// go into subfolders but not archives or links to folders
				CComPtr<IShellFolder> pChild;
				if (SUCCEEDED(pFolder->BindToObject(pidl,NULL,IID_IShellFolder,(void**)&pChild)))
					LoadFolderIcons(pChild,level+1);
			}
		}
		ILFree(pidl);
		if (s_bStopLoading) break;
	}
}

static KNOWNFOLDERID g_CacheFolders[]=
{
	FOLDERID_StartMenu,
	FOLDERID_CommonStartMenu,
	FOLDERID_ControlPanelFolder,
	FOLDERID_Favorites,
};

DWORD CALLBACK CIconManager::PreloadThread( void *param )
{
	// wait 5 seconds before starting the preloading
	// so we don't interfere with the other boot-time processes
	for (int i=0;i<50;i++)
	{
		if (s_bStopLoading) return 0;
		Sleep(100);
	}

	wchar_t path[_MAX_PATH];
	GetModuleFileName(g_Instance,path,_countof(path));
	LoadLibrary(path); // stop the DLL from unloading
	SetThreadPriority(GetCurrentThread(),THREAD_PRIORITY_IDLE);
	CoInitialize(NULL);
	CComPtr<IShellFolder> pDesktop;
	SHGetDesktopFolder(&pDesktop);
	for (int i=0;i<_countof(g_CacheFolders);i++)
	{
		PIDLIST_ABSOLUTE path;
		SHGetKnownFolderIDList(g_CacheFolders[i],0,NULL,&path);
		CComPtr<IShellFolder> pFolder;
		pDesktop->BindToObject(path,NULL,IID_IShellFolder,(void**)&pFolder);
		LoadFolderIcons(pFolder,g_CacheFolders[i]==FOLDERID_ControlPanelFolder?MAX_FOLDER_LEVEL:0);
		ILFree(path);
		if (s_bStopLoading) break;
	}
	CoUninitialize();
	FreeLibraryAndExitThread(g_Instance,0); // release the DLL
}
