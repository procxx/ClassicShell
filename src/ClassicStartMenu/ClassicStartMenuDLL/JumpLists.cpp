// Classic Shell (c) 2009-2012, Ivo Beltchev
// The sources for Classic Shell are distributed under the MIT open source license

#include "stdafx.h"
#include "JumpLists.h"
#include "ResourceHelper.h"
#include "Translations.h"
#include "FNVHash.h"
#include "LogManager.h"
#include <propkey.h>
#include <StrSafe.h>

static KNOWNFOLDERID g_KnownPrefixes[]=
{
	FOLDERID_SystemX86,
	FOLDERID_System,
	FOLDERID_Windows,
	FOLDERID_ProgramFilesX86,
	FOLDERID_ProgramFilesX64,
};

struct DestListHeader
{
	int type; // 1
	int count;
	int pinCount;
	int reserved1;
	int lastStream;
	int reserved2;
	int writeCount;
	int reserved3;
};

struct DestListItemHeader
{
	unsigned __int64 crc;
	char pad1[80];
	int stream;
	int pad2;
	float useCount;
	FILETIME timestamp;
	int pinIdx;
};

struct DestListItem
{
	DestListItemHeader header;
	CString name;
};

struct CustomListHeader
{
	int type; // 2
	int groupCount;
	int reserved;
};

// app ID resolver interface as described here: http://www.binrand.com/post/1510934-out-using-system-using-system-collections-generic-using-system.html
interface IApplicationResolver: public IUnknown
{
	STDMETHOD(GetAppIDForShortcut)( IShellItem *psi, LPWSTR *ppszAppID );
	// .... we don't care about the rest of the methods ....
};

GUID CLSID_ApplicationResolver={0x660b90c8,0x73a9,0x4b58,{0x8c,0xae,0x35,0x5b,0x7f,0x55,0x34,0x1b}};
// different IIDs for Win8 and Win8: http://a-whiter.livejournal.com/1266.html
GUID IID_IApplicationResolver7={0x46a6eeff,0x908e,0x4dc6,{0x92,0xa6,0x64,0xbe,0x91,0x77,0xb4,0x1c}};
GUID IID_IApplicationResolver8={0xde25675a,0x72de,0x44b4,{0x93,0x73,0x05,0x17,0x04,0x50,0xc1,0x40}};

static CComPtr<IApplicationResolver> g_pAppResolver;
static int g_AppResolverTime;

// Creates the app id resolver object
void CreateAppResolver( void )
{
	if (GetWinVersion()>=WIN_VER_WIN7)
	{
		int time=GetTickCount();
		if (!g_pAppResolver || (time-g_AppResolverTime)>60000)
		{
			// recreate the app resolver at most once per minute, as it may need to read lots of data from disk
			g_AppResolverTime=time;
			CComPtr<IUnknown> pUnknown;
			pUnknown.CoCreateInstance(CLSID_ApplicationResolver);
			if (GetWinVersion()==WIN_VER_WIN7)
				g_pAppResolver=CComQIPtr<IApplicationResolver,&IID_IApplicationResolver7>(pUnknown);
			else
				g_pAppResolver=CComQIPtr<IApplicationResolver,&IID_IApplicationResolver8>(pUnknown);
		}
	}
}

// Returns the App ID and the target exe for the given shortcut
// appid must be _MAX_PATH characters
// exe must be _MAX_PATH characters (optional)
bool GetAppInfoForLink( PIDLIST_ABSOLUTE pidl, wchar_t *appid, wchar_t *appexe )
{
	if (GetWinVersion()<WIN_VER_WIN7)
		return false;

	CComPtr<IShellFolder> pFolder;
	PCUITEMID_CHILD child;
	if (FAILED(SHBindToParent(pidl,IID_IShellFolder,(void**)&pFolder,&child)))
		return false;

	CComPtr<IShellLink> pLink;
	if (FAILED(pFolder->GetUIObjectOf(NULL,1,&child,IID_IShellLink,NULL,(void**)&pLink)))
		return false;

	if (appexe)
	{
		if (FAILED(pLink->Resolve(NULL,SLR_INVOKE_MSI|SLR_NO_UI|SLR_NOUPDATE)))
			return false;

		PIDLIST_ABSOLUTE target;
		if (FAILED(pLink->GetIDList(&target)))
			return false;

		wchar_t exe[_MAX_PATH];
		HRESULT hr=SHGetPathFromIDList(target,exe);
		ILFree(target);
		if (FAILED(hr))
			return false;
		Strcpy(appexe,_MAX_PATH,exe);
	}

	CComQIPtr<IPropertyStore> pStore=pLink;
	if (pStore)
	{
		// handle explicit appid
		PROPVARIANT val;
		PropVariantInit(&val);
		if (SUCCEEDED(pStore->GetValue(PKEY_AppUserModel_PreventPinning,&val)) && val.vt==VT_BOOL && val.boolVal)
		{
			PropVariantClear(&val);
			return false;
		}
		PropVariantClear(&val);
		if (SUCCEEDED(pStore->GetValue(PKEY_AppUserModel_ID,&val)) && (val.vt==VT_LPWSTR || val.vt==VT_BSTR))
		{
			Strcpy(appid,_MAX_PATH,val.pwszVal);
			PropVariantClear(&val);
			return true;
		}
		PropVariantClear(&val);
	}

	CComPtr<IShellItem> pItem;
	if (FAILED(SHCreateItemFromIDList(pidl,IID_IShellItem,(void**)&pItem)))
		return false;
	wchar_t *pAppId;
	if (FAILED(g_pAppResolver->GetAppIDForShortcut(pItem,&pAppId)))
		return false;
	Strcpy(appid,_MAX_PATH,pAppId);
	CoTaskMemFree(pAppId);
	return true;
}

// 8-byte CRC as described here: http://msdn.microsoft.com/en-us/library/hh554834(v=prot.10).aspx
static unsigned __int64 CalcCRC64( unsigned char *buf, int len, unsigned __int64 crc=0xFFFFFFFFFFFFFFFF )
{
	static unsigned __int64 CRCTable[256];

	if (!CRCTable[1])
	{
		for (int i=0;i<256;i++)
		{
			unsigned __int64 val=i;
			for (int j=0;j<8;j++)
				val=(val>>1)^((val&1)*0x92C64265D32139A4);
			CRCTable[i]=val;
		}
	}

	for (int i=0;i<len;i++)
		crc = CRCTable[(crc^buf[i])&255]^(crc>>8);

	return crc;
}

// Returns true if the given shortcut has a jumplist (it may be empty)
bool HasJumplist( const wchar_t *appid )
{
	ATLASSERT(GetWinVersion()>=WIN_VER_WIN7);
	// the jumplist is stored in a file in the CustomDestinations folder as described here:
	// http://www.4n6k.com/2011/09/jump-list-forensics-appids-part-1.html
	wchar_t APPID[_MAX_PATH];
	Strcpy(APPID,_countof(APPID),appid);
	CharUpper(APPID);
	unsigned __int64 crc=CalcCRC64((unsigned char*)APPID,Strlen(APPID)*2);

	wchar_t *pRecent;
	if (FAILED(SHGetKnownFolderPath(FOLDERID_Recent,0,NULL,&pRecent)))
		return false;

	wchar_t appkey[100];
	Sprintf(appkey,_countof(appkey),L"%I64x",crc);
	wchar_t path1[_MAX_PATH];
	wchar_t path2[_MAX_PATH];
	LOG_MENU(LOG_OPEN,L"Jumplist Check: appid=%s, appkey=%s",appid,appkey);
	Sprintf(path1,_countof(path1),L"%s\\CustomDestinations\\%s.customDestinations-ms",pRecent,appkey);
	Sprintf(path2,_countof(path2),L"%s\\AutomaticDestinations\\%s.automaticDestinations-ms",pRecent,appkey);
	CoTaskMemFree(pRecent);
	return (GetFileAttributes(path1)!=INVALID_FILE_ATTRIBUTES || GetFileAttributes(path2)!=INVALID_FILE_ATTRIBUTES);
}

static void AddJumpItem( CJumpGroup &group, IUnknown *pUnknown )
{
	CJumpItem item;
	item.type=CJumpItem::TYPE_UNKNOWN;
	item.pItem=pUnknown;
	item.hash=0;
	item.bHidden=false;
	item.bHasArguments=false;
	CComQIPtr<IShellItem> pItem=pUnknown;
	if (pItem)
	{
		item.type=CJumpItem::TYPE_ITEM;
		wchar_t *pName;
		if (FAILED(pItem->GetDisplayName(SIGDN_PARENTRELATIVEEDITING,&pName)))
			return;
		item.name=pName;
		CoTaskMemFree(pName);
		if (SUCCEEDED(pItem->GetDisplayName(SIGDN_DESKTOPABSOLUTEPARSING,&pName)))
		{
			LOG_MENU(LOG_OPEN,L"Jumplist Item Path: %s",pName);
			CharUpper(pName);
			item.hash=CalcFNVHash(pName);
			CoTaskMemFree(pName);
		}
		LOG_MENU(LOG_OPEN,L"Jumplist Item Name: %s",item.name);
		group.items.push_back(item);
		return;
	}

	CComQIPtr<IShellLink> pLink=pUnknown;
	if (pLink)
	{
		item.type=CJumpItem::TYPE_LINK;
		CComQIPtr<IPropertyStore> pStore=pLink;
		if (pStore)
		{
			PROPVARIANT val;
			PropVariantInit(&val);
			if (group.type==CJumpGroup::TYPE_TASKS && SUCCEEDED(pStore->GetValue(PKEY_AppUserModel_IsDestListSeparator,&val)) && val.vt==VT_BOOL && val.boolVal)
			{
				item.type=CJumpItem::TYPE_SEPARATOR;
				PropVariantClear(&val);
			}
			else
			{
				if (SUCCEEDED(pStore->GetValue(PKEY_Title,&val)) && (val.vt==VT_LPWSTR || val.vt==VT_BSTR))
				{
					wchar_t name[256];
					SHLoadIndirectString(val.pwszVal,name,_countof(name),NULL);
					item.name=name;
				}
				PropVariantClear(&val);
			}
		}
		PIDLIST_ABSOLUTE pidl;
		if (SUCCEEDED(pLink->GetIDList(&pidl)))
		{
			wchar_t *pName;
			if (item.name.IsEmpty())
			{
				if (SUCCEEDED(SHGetNameFromIDList(pidl,SIGDN_NORMALDISPLAY,&pName)))
				{
					item.name=pName;
					CoTaskMemFree(pName);
				}
			}
			if (SUCCEEDED(SHGetNameFromIDList(pidl,SIGDN_DESKTOPABSOLUTEPARSING,&pName)))
			{
				LOG_MENU(LOG_OPEN,L"Jumplist Link Path: %s",pName);
				CharUpper(pName);
				item.hash=CalcFNVHash(pName);
				CoTaskMemFree(pName);
			}
			ILFree(pidl);
			wchar_t args[1024];
			if (SUCCEEDED(pLink->GetArguments(args,_countof(args))) && args[0])
			{
				LOG_MENU(LOG_OPEN,L"Jumplist Link Args: %s",args);
				item.hash=CalcFNVHash(args,item.hash);
				item.bHasArguments=true;
			}
		}
		LOG_MENU(LOG_OPEN,L"Jumplist Link Name: %s",item.name);
		if (!item.name.IsEmpty())
			group.items.push_back(item);
		return;
	}
}

static unsigned int CalcLinkStreamHash( IStorage *pStorage, int stream )
{
	CComPtr<IShellLink> pLink;
	if (FAILED(pLink.CoCreateInstance(CLSID_ShellLink)))
		return 0;

	{
		wchar_t streamName[100];
		Sprintf(streamName,_countof(streamName),L"%X",stream);
		CComPtr<IStream> pStream;
		if (FAILED(pStorage->OpenStream(streamName,NULL,STGM_READ|STGM_SHARE_EXCLUSIVE,0,&pStream)))
			return 0;
		{
			CComQIPtr<IPersistStream> pPersist=pLink;
			if (!pPersist || FAILED(pPersist->Load(pStream)))
				return 0;
		}
	}

	PIDLIST_ABSOLUTE pidl;
	if (FAILED(pLink->GetIDList(&pidl)))
		return 0;

	unsigned int hash;
	wchar_t *pName;
	if (SUCCEEDED(SHGetNameFromIDList(pidl,SIGDN_DESKTOPABSOLUTEPARSING,&pName)))
	{
		CharUpper(pName);
		hash=CalcFNVHash(pName);
		CoTaskMemFree(pName);
	}
	ILFree(pidl);
	wchar_t args[1024];
	if (SUCCEEDED(pLink->GetArguments(args,_countof(args))))
		hash=CalcFNVHash(args,hash);
	return hash;
}

static void GetKnownCategory( const wchar_t *appid, CJumpGroup &group, APPDOCLISTTYPE type )
{
	CComPtr<IApplicationDocumentLists> pDocList;
	if (SUCCEEDED(pDocList.CoCreateInstance(CLSID_ApplicationDocumentLists)))
	{
		pDocList->SetAppID(appid);
		CComPtr<IObjectArray> pArray;

		if (SUCCEEDED(pDocList->GetList(type,0,IID_IObjectArray,(void**)&pArray)))
		{
			UINT count;
			pArray->GetCount(&count);
			for (UINT i=0;i<count;i++)
			{
				CComPtr<IUnknown> pUnknown;
				if (SUCCEEDED(pArray->GetAt(i,IID_IUnknown,(void**)&pUnknown)))
				{
					AddJumpItem(group,pUnknown);
				}
			}
		}
	}
}

// Returns the jumplist for the given shortcut
bool GetJumplist( const wchar_t *appid, CJumpList &list, int maxCount )
{
	ATLASSERT(GetWinVersion()>=WIN_VER_WIN7);
	list.Clear();

	wchar_t APPID[_MAX_PATH];
	Strcpy(APPID,_countof(APPID),appid);
	CharUpper(APPID);
	unsigned __int64 crc=CalcCRC64((unsigned char*)APPID,Strlen(APPID)*2);

	wchar_t *pRecent;
	if (FAILED(SHGetKnownFolderPath(FOLDERID_Recent,0,NULL,&pRecent)))
		return false;

	wchar_t appkey[100];
	Sprintf(appkey,_countof(appkey),L"%I64x",crc);
	LOG_MENU(LOG_OPEN,L"Get Jumplist: appid=%s, appkey=%s, max=%d",appid,appkey,maxCount);
	wchar_t path1[_MAX_PATH];
	wchar_t path2[_MAX_PATH];
	Sprintf(path1,_countof(path1),L"%s\\CustomDestinations\\%s.customDestinations-ms",pRecent,appkey);
	Sprintf(path2,_countof(path2),L"%s\\AutomaticDestinations\\%s.automaticDestinations-ms",pRecent,appkey);
	CoTaskMemFree(pRecent);

	CComPtr<IStream> pStream;
	if (SUCCEEDED(SHCreateStreamOnFile(path1,STGM_READ,&pStream)))
	{
		CustomListHeader header;
		if (FAILED(pStream->Read(&header,sizeof(header),NULL)))
			return false;

		list.reserved=header.reserved;
		list.groups.resize(header.groupCount+1);
		bool bHasTasks=false;
		for (int groupIdx=1;groupIdx<=header.groupCount;groupIdx++)
		{
			int type;
			if (FAILED(pStream->Read(&type,4,NULL)))
				return false;
			CJumpGroup *pGroup=&list.groups[groupIdx];
			if (type==1)
			{
				// known category
				if (FAILED(pStream->Read(&type,4,NULL)))
					return false;
				if (type==1)
				{
					pGroup->type=CJumpGroup::TYPE_FREQUENT;
					pGroup->name=FindTranslation(L"JumpList.Frequent",L"Frequent");
					GetKnownCategory(appid,*pGroup,ADLT_FREQUENT);
				}
				else if (type==2)
				{
					pGroup->type=CJumpGroup::TYPE_RECENT;
					pGroup->name=FindTranslation(L"JumpList.Recent",L"Recent");
					GetKnownCategory(appid,*pGroup,ADLT_RECENT);
				}
			}
			else
			{
				if (type==0)
				{
					short len;
					wchar_t str[256];
					if (FAILED(pStream->Read(&len,2,NULL)) || FAILED(pStream->Read(str,len*2,NULL)))
						return false;
					str[len]=0;
					pGroup->name0=str;
					wchar_t name[256];
					SHLoadIndirectString(str,name,_countof(name),NULL);
					pGroup->name=name;
					pGroup->type=CJumpGroup::TYPE_CUSTOM;
				}
				else
				{
					if (!bHasTasks)
					{
						// make sure the Tasks group is last
						bHasTasks=true;
						header.groupCount--;
						groupIdx--;
						pGroup=&list.groups[header.groupCount+1];
					}
					pGroup->name=FindTranslation(L"JumpList.Tasks",L"Tasks");
					pGroup->type=CJumpGroup::TYPE_TASKS;
				}
				int count;
				if (FAILED(pStream->Read(&count,4,NULL)))
					return false;
				for (int i=0;i<count;i++)
				{
					GUID clsid;
					if (FAILED(pStream->Read(&clsid,sizeof(clsid),NULL)))
						return false;
					CComPtr<IPersistStream> pPersist;
					HRESULT hr=pPersist.CoCreateInstance(clsid);
					if (FAILED(hr) || FAILED(pPersist->Load(pStream)))
						return false;
					AddJumpItem(*pGroup,pPersist);
				}
			}
			pGroup->bHidden=false;
			DWORD cookie;
			if (FAILED(pStream->Read(&cookie,4,NULL)) || cookie!=0xBABFFBAB)
				return false;
		}
	}
	else
	{
		list.groups.resize(2);
		CJumpGroup &group=list.groups[1];
		group.type=CJumpGroup::TYPE_RECENT;
		group.name=FindTranslation(L"JumpList.Recent",L"Recent");
		GetKnownCategory(appid,group,ADLT_RECENT);
	}

	{
		// update pinned items
		CJumpGroup &group=list.groups[0];
		group.type=CJumpGroup::TYPE_PINNED;
		group.name=FindTranslation(L"JumpList.Pinned",L"Pinned");

		// read the DestList stream as described here: http://www.forensicswiki.org/wiki/Jump_Lists
		std::vector<int> pinStreams;
		CComPtr<IStorage> pStorage;
		if (SUCCEEDED(StgOpenStorageEx(path2,STGM_READ|STGM_TRANSACTED,STGFMT_STORAGE,0,NULL,0,IID_IStorage,(void**)&pStorage)))
		{
			CComPtr<IStream> pStream;
			if (SUCCEEDED(pStorage->OpenStream(L"DestList",NULL,STGM_READ|STGM_SHARE_EXCLUSIVE,0,&pStream)))
			{
				DestListHeader header;
				if (SUCCEEDED(pStream->Read(&header,sizeof(header),NULL)))
				{
					CJumpGroup &group=list.groups[0];
					pinStreams.resize(header.pinCount,-1);
					for (int i=0;i<header.count;i++)
					{
						DestListItemHeader itemHeader;
						if (FAILED(pStream->Read(&itemHeader,sizeof(itemHeader),NULL)))
							break;
						unsigned __int64 crc=itemHeader.crc;
						itemHeader.crc=0;
						if (CalcCRC64((unsigned char*)&itemHeader,sizeof(itemHeader))!=crc)
							break;
						short len;
						if (FAILED(pStream->Read(&len,2,NULL)))
							break;
						LARGE_INTEGER seek={len*2,0};
						if (FAILED(pStream->Seek(seek,STREAM_SEEK_CUR,NULL)))
							break;
						if (itemHeader.pinIdx>=0 && itemHeader.pinIdx<header.pinCount)
							pinStreams[itemHeader.pinIdx]=itemHeader.stream;
					}
				}
			}
		}

		// read pinned streams
		for (std::vector<int>::const_iterator it=pinStreams.begin();it!=pinStreams.end();++it)
			if (*it>=0)
			{
				wchar_t streamName[100];
				Sprintf(streamName,_countof(streamName),L"%X",*it);
				CComPtr<IStream> pStream;
				if (SUCCEEDED(pStorage->OpenStream(streamName,NULL,STGM_READ|STGM_SHARE_EXCLUSIVE,0,&pStream)))
				{
					CComQIPtr<IPersistStream> pPersist;
					if (SUCCEEDED(pPersist.CoCreateInstance(CLSID_ShellLink)) && SUCCEEDED(pPersist->Load(pStream)))
						AddJumpItem(group,pPersist);
				}
			}

		// remove pinned items from the other groups
		for (std::vector<CJumpItem>::iterator it=group.items.begin();it!=group.items.end();++it)
		{
			unsigned int hash=it->hash;
			bool bReplaced=false;
			for (int g=1;g<(int)list.groups.size();g++)
			{
				if (hash==0) break;
				CJumpGroup &group=list.groups[g];
				for (int i=0;i<(int)group.items.size();i++)
				{
					CJumpItem &item=group.items[i];
					if (item.hash==hash)
					{
						if (!bReplaced)
						{
							// replace the pinned item with the found item. there is a better chance for it to be valid
							// for example Chrome's pinned links may have expired icons, but the custom category links have valid icons
							*it=item;
							bReplaced=true;
						}
						item.bHidden=true;
					}
				}
			}
		}
	}

	// limit the item count (not tasks or pinned)
	for (std::vector<CJumpGroup>::iterator it=list.groups.begin();it!=list.groups.end();++it)
	{
		CJumpGroup &group=*it;
		if (group.type!=CJumpGroup::TYPE_TASKS && group.type!=CJumpGroup::TYPE_PINNED)
		{
			for (std::vector<CJumpItem>::iterator it2=group.items.begin();it2!=group.items.end();++it2)
				if (!it2->bHidden)
				{
					it2->bHidden=(maxCount<=0);
					maxCount--;
				}
		}
	}

	// hide empty groups
	for (std::vector<CJumpGroup>::iterator it=list.groups.begin();it!=list.groups.end();++it)
	{
		CJumpGroup &group=*it;
		group.bHidden=true;
		for (std::vector<CJumpItem>::const_iterator it2=group.items.begin();it2!=group.items.end();++it2)
			if (!it2->bHidden)
			{
				group.bHidden=false;
				break;
			}
	}

	return true;
}

// Executes the given item using the correct application
bool ExecuteJumpItem( const wchar_t *appid, const wchar_t *appexe, const CJumpItem &item, HWND hwnd )
{
	ATLASSERT(GetWinVersion()>=WIN_VER_WIN7);
	if (!item.pItem) return false;
	if (item.type==CJumpItem::TYPE_ITEM)
	{
		LOG_MENU(LOG_OPEN,L"Execute Item: name=%s, appid=%s, appexe=%s",item.name,appid,appexe);
		CComQIPtr<IShellItem> pItem=item.pItem;
		if (!pItem)
			return false;
		wchar_t *pName;
		if (FAILED(pItem->GetDisplayName(SIGDN_DESKTOPABSOLUTEPARSING,&pName)))
			return false;
		wchar_t ext[_MAX_EXT];
		Strcpy(ext,_countof(ext),PathFindExtension(pName));
		CoTaskMemFree(pName);

		// find the correct association handler and invoke it on the item
		CComPtr<IEnumAssocHandlers> pEnumHandlers;
		if (ext[0] && SUCCEEDED(SHAssocEnumHandlers(ext,ASSOC_FILTER_RECOMMENDED,&pEnumHandlers)))
		{
			CComPtr<IAssocHandler> pHandler;
			ULONG count;
			while (SUCCEEDED(pEnumHandlers->Next(1,&pHandler,&count)) && count==1)
			{
				CComQIPtr<IObjectWithAppUserModelID> pObject=pHandler;
				bool bFound=false;
				if (pObject)
				{
					wchar_t *pID;
					if (SUCCEEDED(pObject->GetAppID(&pID)))
					{
						// found explicit appid
						bFound=(_wcsicmp(appid,pID)==0);
						CoTaskMemFree(pID);
						if (bFound)
							LOG_MENU(LOG_OPEN,L"Found handler appid");
					}
				}
				if (!bFound)
				{
					wchar_t *pExe;
					if (SUCCEEDED(pHandler->GetName(&pExe)))
					{
						bFound=(_wcsicmp(appexe,pExe)==0);
						CoTaskMemFree(pExe);
						if (bFound)
							LOG_MENU(LOG_OPEN,L"Found handler appexe");
					}
				}
				if (bFound)
				{
					CComPtr<IDataObject> pDataObject;
					if (SUCCEEDED(pItem->BindToHandler(NULL,BHID_DataObject,IID_IDataObject,(void**)&pDataObject)) && SUCCEEDED(pHandler->Invoke(pDataObject)))
						return true;
					break;
				}
				pHandler=NULL;
			}
		}
		// couldn't find a handler, execute the old way
		SHELLEXECUTEINFO execute={sizeof(execute),SEE_MASK_IDLIST|SEE_MASK_FLAG_LOG_USAGE};
		execute.nShow=SW_SHOWNORMAL;
		PIDLIST_ABSOLUTE pidl;
		if (SUCCEEDED(SHGetIDListFromObject(pItem,&pidl)))
		{
			execute.lpIDList=pidl;
			ShellExecuteEx(&execute);
			ILFree(pidl);
		}
		return true;
	}

	if (item.type==CJumpItem::TYPE_LINK)
	{
		// invoke the link through its context menu
		CComQIPtr<IContextMenu> pMenu=item.pItem;
		if (!pMenu) return false;
		HRESULT hr;
		HMENU menu=CreatePopupMenu();
		hr=pMenu->QueryContextMenu(menu,0,1,1000,CMF_DEFAULTONLY);
		if (FAILED(hr))
		{
			DestroyMenu(menu);
			return false;
		}
		int id=GetMenuDefaultItem(menu,FALSE,0);
		if (id>0)
		{
			CMINVOKECOMMANDINFO command={sizeof(command),CMIC_MASK_FLAG_LOG_USAGE};
			command.lpVerb=MAKEINTRESOURCEA(id-1);
			wchar_t path[_MAX_PATH];
			GetModuleFileName(NULL,path,_countof(path));
			if (_wcsicmp(PathFindFileName(path),L"explorer.exe")==0)
				command.fMask|=CMIC_MASK_ASYNCOK;
			command.hwnd=hwnd;
			command.nShow=SW_SHOWNORMAL;
			hr=pMenu->InvokeCommand(&command);
		}
		DestroyMenu(menu);
	}

	return true;
}

// Removes the given item from the jumplist
void RemoveJumpItem( const wchar_t *appid, CJumpList &list, int groupIdx, int itemIdx )
{
	ATLASSERT(GetWinVersion()>=WIN_VER_WIN7);

	CJumpGroup &group=list.groups[groupIdx];
	if (group.type==CJumpGroup::TYPE_FREQUENT || group.type==CJumpGroup::TYPE_RECENT)
	{
		// removing from the standard lists is easy
		CComPtr<IApplicationDestinations> pDestinations;
		if (SUCCEEDED(pDestinations.CoCreateInstance(CLSID_ApplicationDestinations)))
		{
			pDestinations->SetAppID(appid);
			pDestinations->RemoveDestination(group.items[itemIdx].pItem);
		}
		group.items.erase(group.items.begin()+itemIdx);
	}
	else if (group.type==CJumpGroup::TYPE_CUSTOM)
	{
		group.items.erase(group.items.begin()+itemIdx);
		// write out the list
		wchar_t APPID[_MAX_PATH];
		Strcpy(APPID,_countof(APPID),appid);
		CharUpper(APPID);
		unsigned __int64 crc=CalcCRC64((unsigned char*)APPID,Strlen(APPID)*2);

		wchar_t *pRecent;
		if (FAILED(SHGetKnownFolderPath(FOLDERID_Recent,0,NULL,&pRecent)))
			return;

		wchar_t appkey[100];
		Sprintf(appkey,_countof(appkey),L"%I64x",crc);
		wchar_t path1[_MAX_PATH];
		Sprintf(path1,_countof(path1),L"%s\\CustomDestinations\\%s.tmp",pRecent,appkey);
		wchar_t path2[_MAX_PATH];
		Sprintf(path2,_countof(path2),L"%s\\CustomDestinations\\%s.customDestinations-ms",pRecent,appkey);
		CoTaskMemFree(pRecent);

		CComPtr<IStream> pStream;
		if (FAILED(SHCreateStreamOnFile(path1,STGM_WRITE|STGM_CREATE,&pStream)))
			return;
		DWORD groupCount=(int)list.groups.size()-1;
		int type=2;
		CustomListHeader header={2,groupCount,list.reserved};
		if (FAILED(pStream->Write(&header,sizeof(header),NULL)))
			goto err;
		DWORD cookie=0xBABFFBAB;
		// first write tasks
		for (std::vector<CJumpGroup>::const_iterator it=list.groups.begin();it!=list.groups.end();++it)
		{
			const CJumpGroup &group=*it;
			if (group.type!=CJumpGroup::TYPE_TASKS)
				continue;
			DWORD val=2;
			if (FAILED(pStream->Write(&val,4,NULL)))
				goto err;
			val=(int)group.items.size();
			if (FAILED(pStream->Write(&val,4,NULL)))
				goto err;
			for (std::vector<CJumpItem>::const_iterator it2=group.items.begin();it2!=group.items.end();++it2)
			{
				CComQIPtr<IPersistStream> pPersist=it2->pItem;
				CLSID clsid;
				if (!pPersist || FAILED(pPersist->GetClassID(&clsid)))
					goto err;
				if (FAILED(pStream->Write(&clsid,sizeof(clsid),NULL)) || FAILED(pPersist->Save(pStream,TRUE)))
					goto err;
			}
			if (FAILED(pStream->Write(&cookie,4,NULL)))
				goto err;
		}
		// write custom and known groups
		for (std::vector<CJumpGroup>::const_iterator it=list.groups.begin();it!=list.groups.end();++it)
		{
			const CJumpGroup &group=*it;
			DWORD val;
			switch (group.type)
			{
				case CJumpGroup::TYPE_RECENT:
				case CJumpGroup::TYPE_FREQUENT:
					val=1;
					if (FAILED(pStream->Write(&val,4,NULL)))
						goto err;
					val=(group.type==CJumpGroup::TYPE_RECENT)?2:1;
					if (FAILED(pStream->Write(&val,4,NULL)))
						goto err;
					break;
				case CJumpGroup::TYPE_CUSTOM:
					val=0;
					if (FAILED(pStream->Write(&val,4,NULL)))
						goto err;
					val=Strlen(group.name0);
					if (FAILED(pStream->Write(&val,2,NULL)) || FAILED(pStream->Write((const wchar_t*)group.name0,val*2,NULL)))
						goto err;
					val=(int)group.items.size();
					if (FAILED(pStream->Write(&val,4,NULL)))
						goto err;
					for (std::vector<CJumpItem>::const_iterator it2=group.items.begin();it2!=group.items.end();++it2)
					{
						CComQIPtr<IPersistStream> pPersist=it2->pItem;
						CLSID clsid;
						if (!pPersist || FAILED(pPersist->GetClassID(&clsid)))
							goto err;
						if (FAILED(pStream->Write(&clsid,sizeof(clsid),NULL)) || FAILED(pPersist->Save(pStream,TRUE)))
							goto err;
					}
					break;
				case CJumpGroup::TYPE_TASKS:
				case CJumpGroup::TYPE_PINNED:
					continue;
			}
			DWORD cookie=0xBABFFBAB;
			if (FAILED(pStream->Write(&cookie,4,NULL)))
				goto err;
		}
		pStream=NULL;
		if (MoveFileEx(path1,path2,MOVEFILE_REPLACE_EXISTING))
			return;
err:
		DeleteFile(path1);
	}
}

// Pins or unpins the given item from the jumplist
void PinJumpItem( const wchar_t *appid, const CJumpList &list, int groupIdx, int itemIdx, bool bPin )
{
	wchar_t APPID[_MAX_PATH];
	Strcpy(APPID,_countof(APPID),appid);
	CharUpper(APPID);
	unsigned __int64 crc=CalcCRC64((unsigned char*)APPID,Strlen(APPID)*2);

	wchar_t *pRecent;
	if (FAILED(SHGetKnownFolderPath(FOLDERID_Recent,0,NULL,&pRecent)))
		return;

	wchar_t path[_MAX_PATH];
	Sprintf(path,_countof(path),L"%s\\AutomaticDestinations\\%I64x.automaticDestinations-ms",pRecent,crc);
	CoTaskMemFree(pRecent);

	CJumpGroup::Type groupType=list.groups[groupIdx].type;
	const CJumpItem &jumpItem=list.groups[groupIdx].items[itemIdx];

	// open the jumplist file
	CComPtr<IStorage> pStorage;
	if (FAILED(StgOpenStorageEx(path,STGM_READWRITE|STGM_TRANSACTED,STGFMT_STORAGE,0,NULL,0,IID_IStorage,(void**)&pStorage)))
		pStorage=NULL;

	std::vector<DestListItem> items;
	int foundIndex=-1;
	DestListHeader header;

	if (pStorage)
	{
		// read DestList
		CComPtr<IStream> pStream;
		int pinCount=0;
		if (SUCCEEDED(pStorage->OpenStream(L"DestList",NULL,STGM_READ|STGM_SHARE_EXCLUSIVE,0,&pStream)))
		{
			if (FAILED(pStream->Read(&header,sizeof(header),NULL)))
				return;
			for (int i=0;i<header.count;i++)
			{
				DestListItem item;
				if (FAILED(pStream->Read(&item.header,sizeof(item.header),NULL)))
					return;
				short len;
				if (FAILED(pStream->Read(&len,2,NULL)))
					return;
				wchar_t name[1024];
				if (len>=_countof(name)) return;
				if (FAILED(pStream->Read(name,len*2,NULL)))
					return;
				name[len]=0;
				item.name=name;
				items.push_back(item);
				if (item.header.pinIdx>=0)
					pinCount++;
				if (foundIndex==-1)
				{
					if (CalcLinkStreamHash(pStorage,item.header.stream)==jumpItem.hash)
						foundIndex=i;
				}
			}
			if (header.pinCount!=pinCount)
				return;
		}
	}

	bool bNewStorage=false;
	if (groupType==CJumpGroup::TYPE_CUSTOM)
	{
		ATLASSERT(bPin);
		// pin a custom item
		if (foundIndex!=-1)
			return; // already pinned

		// add new named stream
		if (!pStorage)
		{
			// create the file if it doesn't exist
			if (FAILED(StgCreateStorageEx(path,STGM_READWRITE|STGM_CREATE|STGM_TRANSACTED,STGFMT_STORAGE,0,NULL,NULL,IID_IStorage,(void**)&pStorage)))
				return;
			bNewStorage=true;
			memset(&header,0,sizeof(header));
			header.type=1;
		}
		int maxStream=0;
		for (std::vector<DestListItem>::const_iterator it=items.begin();it!=items.end();++it)
		{
			if (maxStream<it->header.stream)
				maxStream=it->header.stream;
		}
		DestListItem item={0};
		item.header.stream=maxStream+1;
		item.header.pinIdx=header.pinCount;
		item.header.crc=CalcCRC64((unsigned char*)&item.header,sizeof(DestListItemHeader));
		wchar_t streamName[100];
		Sprintf(streamName,_countof(streamName),L"%X",item.header.stream);
		CComPtr<IStream> pStream;
		if (FAILED(pStorage->CreateStream(streamName,STGM_WRITE|STGM_CREATE|STGM_SHARE_EXCLUSIVE,0,0,&pStream)))
			goto end;
		CComQIPtr<IShellLink> pLink=jumpItem.pItem;
		if (pLink)
		{
			wchar_t text[INFOTIPSIZE];
			if (pLink->GetPath(text,_countof(text),NULL,SLGP_RAWPATH)==S_OK)
			{
				// for links with a valid path the name is a crc of the path, arguments, and title
				CharUpper(text);
				unsigned __int64 crc=CalcCRC64((unsigned char*)text,Strlen(text)*2);
				if (pLink->GetArguments(text,_countof(text))==S_OK)
					crc=CalcCRC64((unsigned char*)text,Strlen(text)*2,crc);
				CComQIPtr<IPropertyStore> pStore=pLink;
				if (pStore)
				{
					PROPVARIANT val;
					PropVariantInit(&val);
					if (SUCCEEDED(pStore->GetValue(PKEY_Title,&val)) && (val.vt==VT_LPWSTR || val.vt==VT_BSTR))
					{
						SHLoadIndirectString(val.pwszVal,text,_countof(text),NULL);
						CharUpper(text);
						crc=CalcCRC64((unsigned char*)text,Strlen(text)*2,crc);
					}
					PropVariantClear(&val);
				}
				Sprintf(text,_countof(text),L"%I64x",crc);
				item.name=text;
			}
			else
			{
				// for links with no path (like IE) the name is generated from the pidl
				PIDLIST_ABSOLUTE pidl;
				if (FAILED(pLink->GetIDList(&pidl)))
					goto end;

				wchar_t *pName;
				if (SUCCEEDED(SHGetNameFromIDList(pidl,SIGDN_DESKTOPABSOLUTEPARSING,&pName)))
				{
					item.name=pName;
					CoTaskMemFree(pName);
				}
				else
				{
					ILFree(pidl);
					goto end;
				}
			}
		}
		else
		{
			CComQIPtr<IShellItem> pItem=jumpItem.pItem;
			wchar_t *pName;
			if (!pItem || FAILED(pItem->GetDisplayName(SIGDN_DESKTOPABSOLUTEPARSING,&pName)))
				goto end;

			item.name=pName;
			CoTaskMemFree(pName);

			if (FAILED(pLink.CoCreateInstance(CLSID_ShellLink)) || FAILED(pLink->SetPath(pName)))
				goto end;
		}

		CComQIPtr<IPersistStream> pPersist=pLink;
		if (!pPersist || FAILED(pPersist->Save(pStream,FALSE)))
			goto end;
		items.push_back(item);
		header.lastStream=item.header.stream;
		header.count++;
		header.pinCount++;
	}
	else if (groupType==CJumpGroup::TYPE_FREQUENT || groupType==CJumpGroup::TYPE_RECENT)
	{
		ATLASSERT(bPin);
		// pin a standard item (set pinIndex in DestList)
		if (foundIndex==-1)
			return; // not in DestList, bad
		if (items[foundIndex].header.pinIdx>=0)
			return; // already pinned
		items[foundIndex].header.pinIdx=header.pinCount;
		header.pinCount++;
	}
	else if (groupType==CJumpGroup::TYPE_PINNED)
	{
		// unpin
		if (foundIndex==-1)
			return; // not in DestList, bad
		for (std::vector<DestListItem>::iterator it=items.begin();it!=items.end();++it)
			if (it->header.pinIdx>items[foundIndex].header.pinIdx)
				it->header.pinIdx--;
		bool bUsed=(items[foundIndex].header.useCount>0);
		if (bUsed)
		{
			// if recent history is disabled, also consider the item unused
			CRegKey regKey;
			if (regKey.Open(HKEY_CURRENT_USER,L"Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Advanced")==ERROR_SUCCESS)
			{
				DWORD val;
				bUsed=(regKey.QueryDWORDValue(L"Start_TrackDocs",val)!=ERROR_SUCCESS || val);
			}
		}
		if (bUsed)
		{
			// unpin used item - just clear pinIdx
			if (items[foundIndex].header.pinIdx<0)
				return; // already unpinned
			items[foundIndex].header.pinIdx=-1;
		}
		else
		{
			// unpin unused item - delete stream and remove from DestList
			wchar_t streamName[100];
			Sprintf(streamName,_countof(streamName),L"%X",items[foundIndex].header.stream);
			if (FAILED(pStorage->DestroyElement(streamName)))
				return;

			items.erase(items.begin()+foundIndex);
			header.count--;
		}
		header.pinCount--;
	}
	else
		return; // not supported

	// update CRC
	for (std::vector<DestListItem>::iterator it=items.begin();it!=items.end();++it)
	{
		it->header.crc=0;
		it->header.crc=CalcCRC64((unsigned char*)&it->header,sizeof(DestListItemHeader));
	}

	if (pStorage)
	{
		// write DestList
		CComPtr<IStream> pStream;
		if (FAILED(pStorage->CreateStream(L"DestList",STGM_WRITE|STGM_CREATE|STGM_SHARE_EXCLUSIVE,0,0,&pStream)))
			goto end;

		header.writeCount++;
		if (FAILED(pStream->Write(&header,sizeof(header),NULL)))
			goto end;
		for (std::vector<DestListItem>::const_iterator it=items.begin();it!=items.end();++it)
		{
			const DestListItem &item=*it;
			if (FAILED(pStream->Write(&item.header,sizeof(item.header),NULL)))
				goto end;
			short len=(short)Strlen(item.name);
			if (FAILED(pStream->Write(&len,2,NULL)))
				goto end;
			if (FAILED(pStream->Write((const wchar_t *)item.name,len*2,NULL)))
				goto end;
		}

		pStorage->Commit(STGC_DEFAULT);
	}
	return;

end:
	if (bNewStorage)
	{
		pStorage->Revert();
		pStorage=NULL;
		DeleteFile(path);
	}
}
