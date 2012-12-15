// Classic Shell (c) 2009-2012, Ivo Beltchev
// The sources for Classic Shell are distributed under the MIT open source license

#include "stdafx.h"
#include "MetroLinkManager.h"
#include "IconManager.h"
#include "FNVHash.h"
#include <propkey.h>
#include <map>

struct MetroLinkInfo
{
	FILETIME timestamp;
	CString appid;
	CString name;
	CString package;
	CString packagePath;
	CString iconPath;
	DWORD color;
};

static std::map<unsigned int,MetroLinkInfo> g_MetroLinkCache;

// Returns a list of all metro links from all packages
void GetMetroLinks( std::vector<CString> &links )
{
	wchar_t path[_MAX_PATH]=L"%LOCALAPPDATA%\\Microsoft\\Windows\\Application Shortcuts";
	DoEnvironmentSubst(path,_countof(path));

	wchar_t find1[_MAX_PATH];
	Sprintf(find1,_countof(find1),L"%s\\*.*",path);

	// enumerate package folders
	WIN32_FIND_DATA data1;
	HANDLE h1=FindFirstFile(find1,&data1);
	while (h1!=INVALID_HANDLE_VALUE)
	{
		if ((data1.dwFileAttributes&FILE_ATTRIBUTE_DIRECTORY) && wcscmp(data1.cFileName,L".")!=0 && wcscmp(data1.cFileName,L"..")!=0)
		{
			// enumerate links in the package
			wchar_t find2[_MAX_PATH];
			Sprintf(find2,_countof(find2),L"%s\\%s\\*.lnk",path,data1.cFileName);
			WIN32_FIND_DATA data2;
			HANDLE h2=FindFirstFile(find2,&data2);
			while (h2!=INVALID_HANDLE_VALUE)
			{
				CString name;
				name.Format(L"%s\\%s\\%s",path,data1.cFileName,data2.cFileName);
				links.push_back(name);
				if (!FindNextFile(h2,&data2))
				{
					FindClose(h2);
					break;
				}
			}
		}
		if (!FindNextFile(h1,&data1))
		{
			FindClose(h1);
			break;
		}
	}
}

PROPERTYKEY PKEY_MetroLinkFolder={{0x9F4C2855, 0x9F79, 0x4B39, {0xA8, 0xD0, 0xE1, 0xD4, 0x2D, 0xE1, 0xD5, 0xF3}}, 15};
PROPERTYKEY PKEY_MetroLinkPackage={{0x9F4C2855, 0x9F79, 0x4B39, {0xA8, 0xD0, 0xE1, 0xD4, 0x2D, 0xE1, 0xD5, 0xF3}}, 21};
PROPERTYKEY PKEY_MetroLinkIcon={{0x86D40B4D, 0x9069, 0x443C, {0x81, 0x9A, 0x2A, 0x54, 0x09, 0x0D, 0xCC, 0xEC}}, 2};
PROPERTYKEY PKEY_MetroLinkColor={{0x86D40B4D, 0x9069, 0x443C, {0x81, 0x9A, 0x2A, 0x54, 0x09, 0x0D, 0xCC, 0xEC}}, 4};

bool GetMetroLinkInfo( const wchar_t *path, CString *pName, int *pIcon, bool bLarge )
{
	WIN32_FILE_ATTRIBUTE_DATA data;
	if (!GetFileAttributesEx(path,GetFileExInfoStandard,&data))
		return false;

	wchar_t PATH[_MAX_PATH];
	Strcpy(PATH,_countof(PATH),path);
	CharUpper(PATH);
	unsigned int key=CalcFNVHash(PATH);
	{
		std::map<unsigned int,MetroLinkInfo>::const_iterator it=g_MetroLinkCache.find(key);
		if (it!=g_MetroLinkCache.end() && CompareFileTime(&data.ftLastWriteTime,&it->second.timestamp)==0)
		{
			if (it->second.name.IsEmpty())
				return false;
			*pName=it->second.name;
			if (pIcon)
				*pIcon=g_IconManager.GetMetroIcon(it->second.packagePath,it->second.iconPath,it->second.color,bLarge);
			return true;
		}
	}
	MetroLinkInfo &info=g_MetroLinkCache[key];
	info.timestamp=data.ftLastWriteTime;

	CComPtr<IShellLink> pLink;
	CString appid, package, packagePath, iconPath;
	DWORD color;
	if (FAILED(pLink.CoCreateInstance(CLSID_ShellLink)))
		return false;
	CComQIPtr<IPersistFile> pFile=pLink;
	if (!pFile || FAILED(pFile->Load(path,STGM_READ)))
		return false;
	CComQIPtr<IPropertyStore> pStore=pLink;
	if (!pStore)
		return false;

	PROPVARIANT val;
	PropVariantInit(&val);
	if (FAILED(pStore->GetValue(PKEY_AppUserModel_ID,&val)))
		return false;
	if (val.vt==VT_LPWSTR || val.vt==VT_BSTR)
		appid=val.pwszVal;
	PropVariantClear(&val);
	if (appid.IsEmpty())
		return false;
	if (FAILED(pStore->GetValue(PKEY_MetroLinkPackage,&val)))
		return false;
	if (val.vt==VT_LPWSTR || val.vt==VT_BSTR)
		package=val.pwszVal;
	PropVariantClear(&val);
	if (package.IsEmpty())
		return false;
	if (FAILED(pStore->GetValue(PKEY_MetroLinkFolder,&val)))
		return false;
	if (val.vt==VT_LPWSTR || val.vt==VT_BSTR)
	{
		wchar_t str[_MAX_PATH];
		Strcpy(str,_countof(str),val.pwszVal);
		DoEnvironmentSubst(str,_countof(str));
		packagePath=str;
	}
	PropVariantClear(&val);
	if (packagePath.IsEmpty())
		return false;
	if (FAILED(pStore->GetValue(PKEY_MetroLinkIcon,&val)))
		return false;
	if (val.vt==VT_LPWSTR || val.vt==VT_BSTR)
	{
		if (_wcsnicmp(val.pwszVal,L"ms-resource:",12)==0)
			iconPath=val.pwszVal+12;
		else
			iconPath=val.pwszVal;
	}
	PropVariantClear(&val);
	if (iconPath.IsEmpty())
		return false;
	if (FAILED(pStore->GetValue(PKEY_MetroLinkColor,&val)))
		return false;
	if (val.vt==VT_I4 || val.vt==VT_UI4)
		color=val.intVal;
	else
		color=0;
	PropVariantClear(&val);
	// look in the registry for the display name. there must be a better way to get the name based on the link properties
	CRegKey keyApp;
	wchar_t text1[1024];
	Sprintf(text1,_countof(text1),L"Software\\Classes\\Local Settings\\Software\\Microsoft\\Windows\\CurrentVersion\\AppModel\\Repository\\Packages\\%s\\Applications\\%s",package,appid);
	wchar_t text2[1024];
	Sprintf(text2,_countof(text2),L"Software\\Classes\\Local Settings\\Software\\Microsoft\\Windows\\CurrentVersion\\AppModel\\Repository\\Packages\\%s",package);
	wchar_t str[256];
	if (keyApp.Open(HKEY_CURRENT_USER,text1,KEY_READ)==ERROR_SUCCESS || keyApp.Open(HKEY_CURRENT_USER,text2,KEY_READ)==ERROR_SUCCESS)
	{
		ULONG size=_countof(text1);
		if (keyApp.QueryStringValue(L"DisplayName",text1,&size)==ERROR_SUCCESS)
			SHLoadIndirectString(text1,str,_countof(str),NULL);
	}
	else
	{
		CComPtr<IShellItem> pItem;
		if (FAILED(SHCreateItemFromParsingName(path,NULL,IID_IShellItem,(void**)&pItem)))
			return false;
		wchar_t *pName;
		if (FAILED(pItem->GetDisplayName(SIGDN_NORMALDISPLAY,&pName)))
			return false;
		Strcpy(str,_countof(str),pName);
		CoTaskMemFree(pName);
	}
	info.appid=appid;
	info.name=str;
	info.package=package;
	info.packagePath=packagePath;
	info.iconPath=iconPath;
	info.color=color;
	*pName=str;
	if (pIcon)
		*pIcon=g_IconManager.GetMetroIcon(packagePath,iconPath,color,bLarge);
	return true;
}

#ifndef __IApplicationActivationManager_INTERFACE_DEFINED__
#define __IApplicationActivationManager_INTERFACE_DEFINED__

enum ACTIVATEOPTIONS
{
    AO_NONE	= 0,
    AO_DESIGNMODE	= 0x1,
    AO_NOERRORUI	= 0x2,
    AO_NOSPLASHSCREEN	= 0x4
};

MIDL_INTERFACE("2e941141-7f97-4756-ba1d-9decde894a3d")
IApplicationActivationManager: public IUnknown
{
public:
    virtual HRESULT STDMETHODCALLTYPE ActivateApplication( 
        /* [in] */ __RPC__in LPCWSTR appUserModelId,
        /* [unique][in] */ __RPC__in_opt LPCWSTR arguments,
        /* [in] */ ACTIVATEOPTIONS options,
        /* [out] */ __RPC__out DWORD *processId) = 0;
    
    virtual HRESULT STDMETHODCALLTYPE ActivateForFile( 
        /* [in] */ __RPC__in LPCWSTR appUserModelId,
        /* [in] */ __RPC__in_opt IShellItemArray *itemArray,
        /* [unique][in] */ __RPC__in_opt LPCWSTR verb,
        /* [out] */ __RPC__out DWORD *processId) = 0;
    
    virtual HRESULT STDMETHODCALLTYPE ActivateForProtocol( 
        /* [in] */ __RPC__in LPCWSTR appUserModelId,
        /* [in] */ __RPC__in_opt IShellItemArray *itemArray,
        /* [out] */ __RPC__out DWORD *processId) = 0;
    
};

static const CLSID CLSID_ApplicationActivationManager={0x45BA127D,0x10A8,0x46EA,{0x8A,0xB7,0x56,0xEA,0x90,0x78,0x94,0x3C}};

#endif

static DWORD CALLBACK ExecuteThread( void *param )
{
	CoInitialize(NULL);
	CComPtr<IApplicationActivationManager> pActivationManager;
	if (SUCCEEDED(pActivationManager.CoCreateInstance(CLSID_ApplicationActivationManager)))
	{
		DWORD pid;
		pActivationManager->ActivateApplication((wchar_t*)param,NULL,AO_NONE,&pid);
	}
	CoUninitialize();
	free(param);
	return 0;
}

void ExecuteMetroLink( const wchar_t *path )
{
	CString appid;

	wchar_t PATH[_MAX_PATH];
	Strcpy(PATH,_countof(PATH),path);
	CharUpper(PATH);
	unsigned int key=CalcFNVHash(PATH);

	std::map<unsigned int,MetroLinkInfo>::const_iterator it=g_MetroLinkCache.find(key);
	if (it!=g_MetroLinkCache.end())
		appid=it->second.appid;
	else
	{
		CComPtr<IShellLink> pLink;
		if (FAILED(pLink.CoCreateInstance(CLSID_ShellLink)))
			return;
		CComQIPtr<IPersistFile> pFile=pLink;
		if (!pFile || FAILED(pFile->Load(PATH,STGM_READ)))
			return;
		CComQIPtr<IPropertyStore> pStore=pLink;
		if (!pStore)
			return;

		PROPVARIANT val;
		PropVariantInit(&val);
		if (FAILED(pStore->GetValue(PKEY_AppUserModel_ID,&val)))
			return;
		if (val.vt==VT_LPWSTR || val.vt==VT_BSTR)
			appid=val.pwszVal;
		PropVariantClear(&val);
	}
	if (!appid.IsEmpty())
	{
		wchar_t exePath[_MAX_PATH];
		GetModuleFileName(NULL,exePath,_countof(exePath));
		if (_wcsicmp(PathFindFileName(exePath),L"explorer.exe")==0)
			CreateThread(NULL,0,ExecuteThread,_wcsdup(appid),0,NULL);
		else
			ExecuteThread(_wcsdup(appid));
	}
}
