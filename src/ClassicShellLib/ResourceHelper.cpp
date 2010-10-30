// Classic Shell (c) 2009-2010, Ivo Beltchev
// The sources for Classic Shell are distributed under the MIT open source license

#include <windows.h>
#include "StringSet.h"
#include "StringUtils.h"
#include "Translations.h"
#include <shlobj.h>
#include <vector>

static HINSTANCE g_MainInstance;
static CStringSet g_ResStrings;
static std::map<int,std::vector<char> > g_ResDialogs;

// Loads all strings from hLngInstance
// pDialogs is a NULL-terminated list of dialog IDs. They are loaded from hLngInstance if possible, otherwise from hMainInstance
void LoadTranslationResources( HINSTANCE hMainInstance, HINSTANCE hLngInstance, int *pDialogs )
{
	g_MainInstance=hMainInstance;
	if (hLngInstance)
	{
		LoadTranslationOverrides(hLngInstance);
		g_ResStrings.Init(hLngInstance);
	}
	for (int i=0;pDialogs[i];i++)
	{
		int id=pDialogs[i];
		HINSTANCE hInst=hLngInstance;
		HRSRC hrsrc=NULL;
		if (hLngInstance)
			hrsrc=FindResource(hInst,MAKEINTRESOURCE(id),RT_DIALOG);
		if (!hrsrc)
		{
			hInst=hMainInstance;
			hrsrc=FindResource(hInst,MAKEINTRESOURCE(id),RT_DIALOG);
		}
		if (hrsrc)
		{
			HGLOBAL hglb=LoadResource(hInst,hrsrc);
			if (hglb)
			{
				// finally lock the resource
				LPVOID res=LockResource(hglb);
				std::vector<char> &dlg=g_ResDialogs[id];
				dlg.resize(SizeofResource(hInst,hrsrc));
				if (!dlg.empty())
					memcpy(&dlg[0],res,dlg.size());
			}
		}
	}
}

// Returns a localized string
CString LoadStringEx( int stringID )
{
	CString str=g_ResStrings.GetString(stringID);
	if (str.IsEmpty())
		str.LoadString(g_MainInstance,stringID);
	return str;
}

// Returns a localized dialog template
DLGTEMPLATE *LoadDialogEx( int dlgID )
{
	std::map<int,std::vector<char> >::iterator it=g_ResDialogs.find(dlgID);
	if (it==g_ResDialogs.end())
		return NULL;
	if (it->second.empty())
		return NULL;
	return (DLGTEMPLATE*)&it->second[0];
}

// Loads an icon. path can be a path to .ico file, or in the format "module.dll, number"
HICON LoadIcon( int iconSize, const wchar_t *path, std::vector<HMODULE> &modules )
{
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
			return (HICON)LoadImage(_AtlBaseModule.GetResourceInstance(),res,IMAGE_ICON,iconSize,iconSize,LR_DEFAULTCOLOR);
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

HICON LoadIcon( int iconSize, PIDLIST_ABSOLUTE pidl )
{
	HICON hIcon=NULL;
	CComPtr<IShellFolder> pFolder;
	PCITEMID_CHILD child;
	if (SUCCEEDED(SHBindToParent(pidl,IID_IShellFolder,(void**)&pFolder,&child)))
	{
		bool bLarge=(iconSize>GetSystemMetrics(SM_CXSMICON));
		LONG lSize;
		if (bLarge)
			lSize=MAKELONG(iconSize,GetSystemMetrics(SM_CXSMICON));
		else
			lSize=MAKELONG(GetSystemMetrics(SM_CXICON),iconSize);
		CComPtr<IExtractIcon> pExtract;
		if (SUCCEEDED(pFolder->GetUIObjectOf(NULL,1,&child,IID_IExtractIcon,NULL,(void**)&pExtract)))
		{
			// get the icon location
			wchar_t location[_MAX_PATH];
			int index=0;
			UINT flags=0;
			if (pExtract->GetIconLocation(0,location,_countof(location),&index,&flags)==S_OK)
			{
				if (flags&GIL_NOTFILENAME)
				{
					// extract the icon
					HICON hIcon2=NULL;
					HRESULT hr=pExtract->Extract(location,index,bLarge?&hIcon:&hIcon2,bLarge?&hIcon2:&hIcon,lSize);
					if (FAILED(hr))
						hIcon=hIcon2=NULL;
					if (hr==S_FALSE)
						flags=0;
					if (hIcon2) DestroyIcon(hIcon2); // HACK!!! Even though Extract should support NULL, not all implementations do. For example shfusion.dll crashes
				}
				if (!(flags&GIL_NOTFILENAME))
				{
					if (ExtractIconEx(location,index==-1?0:index,bLarge?&hIcon:NULL,bLarge?NULL:&hIcon,1)!=1)
						hIcon=NULL;
				}
			}
		}
		else
		{
			// try again using the ANSI version
			CComPtr<IExtractIconA> pExtractA;
			if (SUCCEEDED(pFolder->GetUIObjectOf(NULL,1,&child,IID_IExtractIconA,NULL,(void**)&pExtractA)))
			{
				// get the icon location
				char location[_MAX_PATH];
				int index=0;
				UINT flags=0;
				if (pExtractA->GetIconLocation(0,location,_countof(location),&index,&flags)==S_OK)
				{
					if (flags&GIL_NOTFILENAME)
					{
						// extract the icon
						HICON hIcon2=NULL;
						HRESULT hr=pExtractA->Extract(location,index,bLarge?&hIcon:&hIcon2,bLarge?&hIcon2:&hIcon,lSize);
						if (FAILED(hr))
							hIcon=hIcon2=NULL;
						if (hr==S_FALSE)
							flags=0;
						if (hIcon2) DestroyIcon(hIcon2); // HACK!!! Even though Extract should support NULL, not all implementations do. For example shfusion.dll crashes
					}
					if (!(flags&GIL_NOTFILENAME))
					{
						if (ExtractIconExA(location,index==-1?0:index,bLarge?&hIcon:NULL,bLarge?NULL:&hIcon,1)!=1)
							hIcon=NULL;
					}
				}
			}
		}
	}

	return hIcon;
}

// Converts an icon to a bitmap. pBits may be NULL. If bDestroyIcon is true, hIcon will be destroyed
HBITMAP BitmapFromIcon( HICON hIcon, int iconSize, unsigned int **pBits, bool bDestroyIcon )
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
	DrawIconEx(hdc,0,0,hIcon,iconSize,iconSize,0,NULL,DI_NORMAL);
	SelectObject(hdc,bmp0);
	DeleteDC(hdc);
	if (bDestroyIcon) DestroyIcon(hIcon);
	if (pBits) *pBits=bits;
	return bmp;
}

// Creates a grayscale version of an icon
HICON CreateDisabledIcon( HICON hIcon, int iconSize )
{
	// convert normal icon to grayscale
	ICONINFO info;
	GetIconInfo(hIcon,&info);

	unsigned int *bits;
	HBITMAP bmp=BitmapFromIcon(hIcon,iconSize,&bits,false);

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
	hIcon=CreateIconIndirect(&info);
	DeleteObject(bmp);
	if (info.hbmMask) DeleteObject(info.hbmMask);
	return hIcon;
}

// Returns the version of a given module
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
