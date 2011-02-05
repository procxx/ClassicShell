// Classic Shell (c) 2009-2010, Ivo Beltchev
// The sources for Classic Shell are distributed under the MIT open source license

#include "stdafx.h"
#include "IconManager.h"
#include "FNVHash.h"
#include "Settings.h"
#include "Translations.h"
#include "ResourceHelper.h"
#include "MenuContainer.h"

const int MAX_FOLDER_LEVEL=10; // don't go more than 10 levels deep

int CIconManager::s_DPI;
int CIconManager::LARGE_ICON_SIZE;
int CIconManager::SMALL_ICON_SIZE;
volatile CIconManager::TLoadingStage CIconManager::s_LoadingStage;
std::map<unsigned int,HICON> CIconManager::s_PreloadedIcons;
std::map<unsigned int,HBITMAP> CIconManager::s_PreloadedBitmaps;
std::vector<CIconManager::IconLocation> CIconManager::s_IconLocations;
CRITICAL_SECTION CIconManager::s_PreloadSection;
CRITICAL_SECTION CIconManager::s_PostloadSection;
HANDLE CIconManager::s_PostloadEvent;

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

	SMALL_ICON_SIZE=GetSettingInt(L"SmallIconSize");
	LARGE_ICON_SIZE=GetSettingInt(L"LargeIconSize");

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
		s_IconLocations.push_back(IconLocation());
		DestroyIcon(info.hIcon);
	}

	s_LoadingStage=LOAD_LOADING;
	InitializeCriticalSection(&s_PreloadSection);
	InitializeCriticalSection(&s_PostloadSection);
	s_PostloadEvent=CreateEvent(NULL,FALSE,FALSE,NULL);

	if (GetSettingBool(L"PreCacheIcons") && _wcsicmp(PathFindFileName(path),L"explorer.exe")==0)
	{
		// don't preload icons if running outside of the explorer
		m_PreloadThread=CreateThread(NULL,0,PreloadThread,NULL,0,NULL);
	}
	else
		m_PreloadThread=INVALID_HANDLE_VALUE;

	if (GetSettingBool(L"DelayIcons"))
		m_PostloadThread=CreateThread(NULL,0,PostloadThread,NULL,0,NULL);
	else
		m_PostloadThread=INVALID_HANDLE_VALUE;
}

void CIconManager::StopLoading( void )
{
	s_LoadingStage=LOAD_STOPPING;
	SetEvent(s_PostloadEvent);
	if (m_PreloadThread!=INVALID_HANDLE_VALUE)
	{
		WaitForSingleObject(m_PreloadThread,INFINITE);
		CloseHandle(m_PreloadThread);
		m_PreloadThread=INVALID_HANDLE_VALUE;
	}
	if (m_PostloadThread!=INVALID_HANDLE_VALUE)
	{
		WaitForSingleObject(m_PostloadThread,INFINITE);
		CloseHandle(m_PostloadThread);
		m_PostloadThread=INVALID_HANDLE_VALUE;
	}
	s_LoadingStage=LOAD_STOPPED;
	DeleteCriticalSection(&s_PostloadSection);
	DeleteCriticalSection(&s_PreloadSection);
}

CIconManager::~CIconManager( void )
{
	if (m_LargeIcons) ImageList_Destroy(m_LargeIcons);
	if (m_SmallIcons) ImageList_Destroy(m_SmallIcons);
}

// Retrieves an icon from a shell folder and child ID
int CIconManager::GetIcon( IShellFolder *pFolder, PCUITEMID_CHILD item, bool bLarge )
{
	ProcessLoadedIcons();

	// get the IExtractIcon object
	CComPtr<IExtractIcon> pExtract;
	HRESULT hr=pFolder->GetUIObjectOf(NULL,1,&item,IID_IExtractIcon,NULL,(void**)&pExtract);
	HICON hIcon=NULL;
	int iconSize=bLarge?LARGE_ICON_SIZE:SMALL_ICON_SIZE;
	bool bUseFactory=false;
	IconLocation loc;
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
		if (flags&GIL_NOTFILENAME)
		{
			HICON hIcon2=NULL;
			HRESULT hr;
			if (iconSize<=GetSystemMetrics(SM_CXSMICON))
				hr=pExtract->Extract(location,index,&hIcon2,&hIcon,MAKELONG(iconSize,iconSize)); // small icon is closer
			else
				hr=pExtract->Extract(location,index,&hIcon,&hIcon2,MAKELONG(iconSize,iconSize)); // large icon is closer
			if (FAILED(hr))
				hIcon=hIcon2=NULL;
			else
			{
				if (hIcon)
				{
					// see if we got the icon with a correct size. use an image factory if the size is too small
					ICONINFO info;
					GetIconInfo(hIcon,&info);
					BITMAP info2;
					GetObject(info.hbmColor,sizeof(info2),&info2);
					if (info.hbmColor) DeleteObject(info.hbmColor);
					if (info.hbmMask) DeleteObject(info.hbmMask);
					if (info2.bmWidth<iconSize)
					{
						DestroyIcon(hIcon);
						hIcon=NULL;
						bUseFactory=true;
					}
				}
				if (hIcon2) DestroyIcon(hIcon2); // HACK!!! Even though Extract should support NULL, not all implementations do. For example shfusion.dll crashes
			}
			if (hr==S_FALSE)
			{
				// we are not supposed to be getting S_FALSE here, but we do (like for EXEs that don't have an icon). fallback to factory
				bUseFactory=true;
			}
		}

		if (!(flags&GIL_NOTFILENAME))
		{
			// the IExtractIcon object didn't do anything - use ShExtractIcon instead
			if (bLarge || m_PostloadThread==INVALID_HANDLE_VALUE)
				hIcon=ShExtractIcon(location,index==-1?0:index,iconSize);
			else
			{
				loc.location=location;
				loc.index=index;
				loc.key=key;
			}
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
		if (flags&GIL_NOTFILENAME)
		{
			HICON hIcon2=NULL;
			HRESULT hr;
			if (iconSize<=GetSystemMetrics(SM_CXSMICON))
				hr=pExtractA->Extract(location,index,&hIcon2,&hIcon,MAKELONG(iconSize,iconSize)); // small icon is closer
			else
				hr=pExtractA->Extract(location,index,&hIcon,&hIcon2,MAKELONG(iconSize,iconSize)); // large icon is closer
			if (FAILED(hr))
				hIcon=hIcon2=NULL;
			else
			{
				if (hIcon)
				{
					// see if we got the icon with a correct size. use an image factory if the size is too small
					ICONINFO info;
					GetIconInfo(hIcon,&info);
					BITMAP info2;
					GetObject(info.hbmColor,sizeof(info2),&info2);
					if (info.hbmColor) DeleteObject(info.hbmColor);
					if (info.hbmMask) DeleteObject(info.hbmMask);
					if (info2.bmWidth<iconSize)
					{
						DestroyIcon(hIcon);
						hIcon=NULL;
						bUseFactory=true;
					}
				}
				if (hIcon2) DestroyIcon(hIcon2); // HACK!!! Even though Extract should support NULL, not all implementations do. For example shfusion.dll crashes
			}
			if (hr==S_FALSE)
			{
				// we are not supposed to be getting S_FALSE here, but we do (like for EXEs that don't have an icon). fallback to factory
				bUseFactory=true;
			}
		}

		if (!(flags&GIL_NOTFILENAME))
		{
			// the IExtractIcon object didn't do anything - use ShExtractIcon instead
			if (bLarge || m_PostloadThread==INVALID_HANDLE_VALUE)
				hIcon=ShExtractIcon(location,index==-1?0:index,iconSize);
			else
			{
				loc.location=location;
				loc.index=index;
				loc.key=key;
			}
		}
	}

	// add to the image list
	int res=0;
	if (bUseFactory)
	{
		// use the image factory to get icons that are larger than the system icon size
		CComPtr<IShellItemImageFactory> pFactory;
		if (SUCCEEDED(SHCreateItemWithParent(NULL,pFolder,item,IID_IShellItemImageFactory,(void**)&pFactory)) && pFactory)
		{
			SIZE size={iconSize,iconSize};
			HBITMAP hBitmap;
			if (SUCCEEDED(pFactory->GetImage(size,SIIGBF_ICONONLY,&hBitmap)))
			{
				res=ImageList_AddMasked(bLarge?m_LargeIcons:m_SmallIcons,hBitmap,CLR_NONE);
				DeleteObject(hBitmap);
			}
		}
		ATLASSERT(!hIcon);
	}
	if (hIcon)
	{
		res=ImageList_AddIcon(bLarge?m_LargeIcons:m_SmallIcons,hIcon);
		DestroyIcon(hIcon);
	}

	if (!loc.location.IsEmpty())
	{
		// post-load
		res=ImageList_GetImageCount(m_SmallIcons);
		ImageList_SetImageCount(m_SmallIcons,res+1);
		ImageList_Copy(m_SmallIcons,res,m_SmallIcons,0,ILCF_MOVE);
		loc.bTemp=true;
	}

	// add to the cache
	if (bLarge)
		m_LargeCache[key]=res;
	else
	{
		EnterCriticalSection(&s_PreloadSection);
		m_SmallCache[key]=res;
		LeaveCriticalSection(&s_PreloadSection);

		if (res>0)
		{
			EnterCriticalSection(&s_PostloadSection);
			s_IconLocations.push_back(loc);
			LeaveCriticalSection(&s_PostloadSection);
		}

#ifdef _DEBUG
		int n=ImageList_GetImageCount(m_SmallIcons);
		ATLASSERT(n==s_IconLocations.size());
#endif
	}

	return res;
}

// Retrieves an icon from a file and icon index (index>=0 - icon index, index<0 - resource ID)
int CIconManager::GetIcon( const wchar_t *location, int index, bool bLarge )
{
	ProcessLoadedIcons();

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
	HMODULE hMod=*location?GetModuleHandle(PathFindFileName(location)):g_Instance;
	int iconSize=bLarge?LARGE_ICON_SIZE:SMALL_ICON_SIZE;
	if (hMod && index<0)
		hIcon=(HICON)LoadImage(hMod,MAKEINTRESOURCE(-index),IMAGE_ICON,iconSize,iconSize,LR_DEFAULTCOLOR);
	else
		hIcon=ShExtractIcon(location,index==-1?0:index,iconSize);

	// add to the image list
	if (hIcon)
	{
		index=ImageList_AddIcon(bLarge?m_LargeIcons:m_SmallIcons,hIcon);
		DestroyIcon(hIcon);
		if (!bLarge)
		{
			EnterCriticalSection(&s_PostloadSection);
			s_IconLocations.push_back(IconLocation());
			LeaveCriticalSection(&s_PostloadSection);
		}
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

#ifdef _DEBUG
		int n=ImageList_GetImageCount(m_SmallIcons);
		ATLASSERT(n==s_IconLocations.size());
#endif
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

void CIconManager::ProcessLoadedIcons( void )
{
	EnterCriticalSection(&s_PreloadSection);
	for (std::map<unsigned int,HICON>::const_iterator it=s_PreloadedIcons.begin();it!=s_PreloadedIcons.end();++it)
	{
		unsigned int key=it->first;
		HICON hIcon=it->second;
		std::map<unsigned int,int>::iterator cache=m_SmallCache.find(key);
		EnterCriticalSection(&s_PostloadSection);
		if (cache==m_SmallCache.end())
		{
			// add to the image list and to the cache
			m_SmallCache[key]=ImageList_AddIcon(m_SmallIcons,hIcon);
			s_IconLocations.push_back(IconLocation());
		}
		else if (s_IconLocations[cache->second].bTemp)
		{
			s_IconLocations[cache->second].bTemp=false;
			ImageList_ReplaceIcon(m_SmallIcons,cache->second,hIcon);
		}
		LeaveCriticalSection(&s_PostloadSection);
		DestroyIcon(hIcon);
	}
	s_PreloadedIcons.clear();
	for (std::map<unsigned int,HBITMAP>::const_iterator it=s_PreloadedBitmaps.begin();it!=s_PreloadedBitmaps.end();++it)
	{
		unsigned int key=it->first;
		HBITMAP hBitmap=it->second;
		EnterCriticalSection(&s_PostloadSection);
		if (m_SmallCache.find(key)==m_SmallCache.end())
		{
			// add to the image list and to the cache
			m_SmallCache[key]=ImageList_AddMasked(m_SmallIcons,hBitmap,CLR_NONE);
			s_IconLocations.push_back(IconLocation());
		}
		else
		{
			ATLASSERT(!s_IconLocations[m_SmallCache[key]].bTemp);
		}
		LeaveCriticalSection(&s_PostloadSection);
		DeleteObject(hBitmap);
	}
	s_PreloadedBitmaps.clear();
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
		unsigned int key;
		bool bUseFactory=false;
		CComPtr<IExtractIcon> pExtract;
		CComPtr<IExtractIconA> pExtractA;
		if (SUCCEEDED(pFolder->GetUIObjectOf(NULL,1,&pidl,IID_IExtractIcon,NULL,(void**)&pExtract)))
		{
			// get the icon location
			wchar_t location[_MAX_PATH];
			int index=0;
			UINT flags=0;
			if (SUCCEEDED(pExtract->GetIconLocation(0,location,_countof(location),&index,&flags)))
			{
				// check if this location+index is in the cache
				key=CalcFNVHash(location,CalcFNVHash(&index,4));
				EnterCriticalSection(&s_PreloadSection);
				bool bLoaded=g_IconManager.m_SmallCache.find(key)!=g_IconManager.m_SmallCache.end() || s_PreloadedIcons.find(key)!=s_PreloadedIcons.end();
				LeaveCriticalSection(&s_PreloadSection);
				if (!bLoaded)
				{
					HICON hIcon, hIcon2=NULL;
					HRESULT hr=E_FAIL;
					if (flags&GIL_NOTFILENAME)
					{
						// extract the icon
						hr=pExtract->Extract(location,index,&hIcon2,&hIcon,MAKELONG(LARGE_ICON_SIZE,SMALL_ICON_SIZE));
						if (hr==S_OK)
						{
							// see if we got the icon with a correct size. use an image factory if the size is too small
							ICONINFO info;
							GetIconInfo(hIcon,&info);
							BITMAP info2;
							GetObject(info.hbmColor,sizeof(info2),&info2);
							if (info.hbmColor) DeleteObject(info.hbmColor);
							if (info.hbmMask) DeleteObject(info.hbmMask);
							if (info2.bmWidth<SMALL_ICON_SIZE)
							{
								DestroyIcon(hIcon);
								hIcon=NULL;
								hr=E_FAIL;
								bUseFactory=true;
							}
							DestroyIcon(hIcon2); // HACK!!! Even though Extract should support NULL, not all implementations do. For example shfusion.dll crashes
						}
						if (hr==S_FALSE)
						{
							// we are not supposed to be getting S_FALSE here, but we do (like for EXEs that don't have an icon). fallback to factory
							bUseFactory=true;
						}
					}
					else
					{
						// the IExtractIcon object didn't do anything - use ShExtractIcon instead
						hIcon=ShExtractIcon(location,index==-1?0:index,SMALL_ICON_SIZE);
						if (hIcon)
						 hr=S_OK;
					}
					if (hr==S_OK)
					{
						EnterCriticalSection(&s_PreloadSection);
						if (s_PreloadedIcons.find(key)!=s_PreloadedIcons.end())
							DestroyIcon(hIcon);
						else
							s_PreloadedIcons[key]=hIcon;
						LeaveCriticalSection(&s_PreloadSection);
					}
					Sleep(10); // pause for a bit to reduce the stress on the system
				}
			}
		}
		else if (SUCCEEDED(pFolder->GetUIObjectOf(NULL,1,&pidl,IID_IExtractIconA,NULL,(void**)&pExtractA))) // try the ANSI version
		{
			// get the icon location
			char location[_MAX_PATH];
			int index=0;
			UINT flags=0;
			if (SUCCEEDED(pExtractA->GetIconLocation(0,location,_countof(location),&index,&flags)))
			{
				// check if this location+index is in the cache
				key=CalcFNVHash(location,CalcFNVHash(&index,4));
				EnterCriticalSection(&s_PreloadSection);
				bool bLoaded=g_IconManager.m_SmallCache.find(key)!=g_IconManager.m_SmallCache.end() || s_PreloadedIcons.find(key)!=s_PreloadedIcons.end();
				LeaveCriticalSection(&s_PreloadSection);
				if (!bLoaded)
				{
					HICON hIcon, hIcon2=NULL;
					HRESULT hr=E_FAIL;
					if (flags&GIL_NOTFILENAME)
					{
						// extract the icon
						hr=pExtractA->Extract(location,index,&hIcon2,&hIcon,MAKELONG(LARGE_ICON_SIZE,SMALL_ICON_SIZE));
						if (hr==S_OK)
						{
							// see if we got the icon with a correct size. use an image factory if the size is too small
							ICONINFO info;
							GetIconInfo(hIcon,&info);
							BITMAP info2;
							GetObject(info.hbmColor,sizeof(info2),&info2);
							if (info.hbmColor) DeleteObject(info.hbmColor);
							if (info.hbmMask) DeleteObject(info.hbmMask);
							if (info2.bmWidth<SMALL_ICON_SIZE)
							{
								DestroyIcon(hIcon);
								hIcon=NULL;
								hr=E_FAIL;
								bUseFactory=true;
							}
							DestroyIcon(hIcon2); // HACK!!! Even though Extract should support NULL, not all implementations do. For example shfusion.dll crashes
						}
						if (hr==S_FALSE)
						{
							// we are not supposed to be getting S_FALSE here, but we do (like for EXEs that don't have an icon). fallback to factory
							bUseFactory=true;
						}
					}
					else
					{
						// the IExtractIcon object didn't do anything - use ShExtractIcon instead
						hIcon=ShExtractIcon(location,index==-1?0:index,SMALL_ICON_SIZE);
						if (hIcon)
							hr=S_OK;
					}
					if (hr==S_OK)
					{
						EnterCriticalSection(&s_PreloadSection);
						if (s_PreloadedIcons.find(key)!=s_PreloadedIcons.end())
							DestroyIcon(hIcon);
						else
							s_PreloadedIcons[key]=hIcon;
						LeaveCriticalSection(&s_PreloadSection);
					}
					Sleep(10); // pause for a bit to reduce the stress on the system
				}
			}
		}

		if (bUseFactory)
		{
			// use the image factory to get icons that are larger than the system icon size
			CComPtr<IShellItemImageFactory> pFactory;
			if (SUCCEEDED(SHCreateItemWithParent(NULL,pFolder,pidl,IID_IShellItemImageFactory,(void**)&pFactory)) && pFactory)
			{
				SIZE size={SMALL_ICON_SIZE,SMALL_ICON_SIZE};
				HBITMAP hBitmap;
				if (SUCCEEDED(pFactory->GetImage(size,SIIGBF_ICONONLY,&hBitmap)))
				{
					EnterCriticalSection(&s_PreloadSection);
					s_PreloadedBitmaps[key]=hBitmap;
					LeaveCriticalSection(&s_PreloadSection);
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
		if (s_LoadingStage!=LOAD_LOADING) break;
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
		if (s_LoadingStage!=LOAD_LOADING) return 0;
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
		if (FAILED(SHGetKnownFolderIDList(g_CacheFolders[i],0,NULL,&path)) || !path) continue;
		CComPtr<IShellFolder> pFolder;
		if (SUCCEEDED(pDesktop->BindToObject(path,NULL,IID_IShellFolder,(void**)&pFolder)) && pFolder)
			LoadFolderIcons(pFolder,g_CacheFolders[i]==FOLDERID_ControlPanelFolder?MAX_FOLDER_LEVEL:0);
		ILFree(path);
		if (s_LoadingStage!=LOAD_LOADING) break;
	}
	CoUninitialize();
	FreeLibraryAndExitThread(g_Instance,0); // release the DLL
}

DWORD CALLBACK CIconManager::PostloadThread( void *param )
{
	// this thread loads any icons marked as bTemp and bUsed
	bool bRefresh=false;
	while (1)
	{
		WaitForSingleObject(s_PostloadEvent,INFINITE);
		bRefresh=false;
		int t0=GetTickCount();
		while (1)
		{
			if (s_LoadingStage!=LOAD_LOADING)
				return 0;
			IconLocation location;
			EnterCriticalSection(&s_PostloadSection);
			for (std::vector<IconLocation>::iterator it=s_IconLocations.begin();it!=s_IconLocations.end();++it)
			{
				if (it->bTemp && !it->bLoaded && it->bUsed)
				{
					it->bLoaded=true;
					location=*it;
					break;
				}
			}
			LeaveCriticalSection(&s_PostloadSection);
			if (location.location.IsEmpty())
				break;
			EnterCriticalSection(&s_PreloadSection);
			bool bLoaded=s_PreloadedIcons.find(location.key)!=s_PreloadedIcons.end(); // check if it was loaded by the pre-load thread
			LeaveCriticalSection(&s_PreloadSection);
			if (!bLoaded)
			{
				HICON hIcon=ShExtractIcon(location.location,location.index==-1?0:location.index,SMALL_ICON_SIZE);
				if (hIcon)
				{
					EnterCriticalSection(&s_PreloadSection);
					if (s_PreloadedIcons.find(location.key)!=s_PreloadedIcons.end())
						DestroyIcon(hIcon);
					else
						s_PreloadedIcons[location.key]=hIcon;
					LeaveCriticalSection(&s_PreloadSection);
				}
			}
			bRefresh=true;
			int t=GetTickCount();
			if (t-t0>100)
			{
				CMenuContainer::RefreshIcons();
				t0=t;
				bRefresh=false;
			}
		}
		if (bRefresh)
			CMenuContainer::RefreshIcons();
	}
}

// Marks all icons as unused (also locks s_PostloadSection)
void CIconManager::ResetUsedIcons( void )
{
	EnterCriticalSection(&s_PostloadSection);
	for (std::vector<IconLocation>::iterator it=s_IconLocations.begin();it!=s_IconLocations.end();++it)
		it->bUsed=false;
}

// Marks the icon as used
void CIconManager::AddUsedIcon( int icon )
{
	s_IconLocations[icon].bUsed=true;
}

// Wakes up the post-loading thread (also unlocks s_PostloadSection)
void CIconManager::StartPostLoading( void )
{
	LeaveCriticalSection(&s_PostloadSection);
	SetEvent(s_PostloadEvent);
}
