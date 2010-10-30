// Classic Shell (c) 2009-2010, Ivo Beltchev
// The sources for Classic Shell are distributed under the MIT open source license

// MenuContainer.cpp - contains the main logic of CMenuContainer

#include "stdafx.h"
#include "IconManager.h"
#include "MenuContainer.h"
#include "Accessibility.h"
#include "ClassicStartMenuDLL.h"
#include "Settings.h"
#include "Translations.h"
#include "CustomMenu.h"
#include "LogManager.h"
#include "FNVHash.h"
#include "ResourceHelper.h"
#include "dllmain.h"
#include "resource.h"
#include <uxtheme.h>
#include <vsstyle.h>
#include <vssym32.h>
#include <WtsApi32.h>
#include <Lm.h>
#include <Dsgetdc.h>
#include <PowrProf.h>
#include <dwmapi.h>
#define SECURITY_WIN32
#include <Security.h>
#include <algorithm>

struct StdMenuOption
{
	TMenuID id;
	int options;
};

// Options for special menu items
enum
{
	MENU_NONE     = 0,
	MENU_ENABLED  = 1, // the item shows in the menu
	MENU_EXPANDED = 2, // the item is expanded
};

static StdMenuOption g_StdOptions[]=
{
	{MENU_FAVORITES,MENU_NONE}, // MENU_ENABLED|MENU_EXPANDED from settings, check policy
	{MENU_DOCUMENTS,MENU_NONE}, // MENU_ENABLED|MENU_EXPANDED from settings, check policy
	{MENU_HELP,MENU_ENABLED}, // check policy
	{MENU_RUN,MENU_ENABLED}, // check policy
	{MENU_LOGOFF,MENU_NONE}, // MENU_ENABLED from settings, check policy
	{MENU_DISCONNECT,MENU_NONE}, // MENU_ENABLED if in a remote session, check policy
	{MENU_SHUTDOWN_BOX,MENU_ENABLED}, // MENU_NONE if in a remote session, check policy
	{MENU_SHUTDOWN,MENU_ENABLED}, // MENU_NONE if in a remote session, check policy
	{MENU_UNDOCK,MENU_ENABLED}, // from settings, check policy
	{MENU_CONTROLPANEL,MENU_ENABLED|MENU_EXPANDED}, // MENU_EXPANDED from settings, check policy
	{MENU_NETWORK,MENU_ENABLED}, // MENU_EXPANDED from settings, check policy
	{MENU_SECURITY,MENU_ENABLED}, // MENU_ENABLED if in a remote session
	{MENU_PRINTERS,MENU_ENABLED}, // MENU_EXPANDED from settings, check policy
	{MENU_TASKBAR,MENU_ENABLED}, // check policy
	{MENU_FEATURES,MENU_ENABLED}, // no setting (prevents the Programs and Features from expanding), check policy (for control panel)
	{MENU_CLASSIC_SETTINGS,MENU_ENABLED}, // MENU_ENABLED from ini file
	{MENU_SEARCH,MENU_ENABLED}, // check policy
	{MENU_SEARCH_FILES,MENU_NONE}, // MENU_ENABLED if the search service is available
	{MENU_SEARCH_PRINTER,MENU_NONE}, // MENU_ENABLED if Active Directory is available
	{MENU_SEARCH_COMPUTERS,MENU_NONE}, // MENU_ENABLED if Active Directory is available, check policy
	{MENU_SEARCH_PEOPLE,MENU_NONE}, // MENU_ENABLED if %ProgramFiles%\Windows Mail\wab.exe exists
	{MENU_USERFILES,MENU_ENABLED}, // check policy
	{MENU_USERDOCUMENTS,MENU_ENABLED}, // check policy
	{MENU_USERPICTURES,MENU_ENABLED}, // check policy
	{MENU_SLEEP,MENU_ENABLED}, // check power caps
	{MENU_HIBERNATE,MENU_ENABLED}, // check power caps
};

///////////////////////////////////////////////////////////////////////////////

int CMenuContainer::s_MaxRecentDocuments=15;
int CMenuContainer::s_ScrollMenus=0;
bool CMenuContainer::s_bRTL=false;
bool CMenuContainer::s_bKeyboardCues=false;
bool CMenuContainer::s_bExpandRight=true;
bool CMenuContainer::s_bRecentItems=false;
bool CMenuContainer::s_bBehindTaskbar=true;
bool CMenuContainer::s_bShowTopEmpty=false;
bool CMenuContainer::s_bNoDragDrop=false;
bool CMenuContainer::s_bNoContextMenu=false;
bool CMenuContainer::s_bExpandLinks=false;
bool CMenuContainer::s_bLogicalSort=false;
bool CMenuContainer::s_bAllPrograms=false;
char CMenuContainer::s_bActiveDirectory=-1;
CMenuContainer *CMenuContainer::s_pDragSource=NULL;
bool CMenuContainer::s_bRightDrag;
std::vector<CMenuContainer*> CMenuContainer::s_Menus;
std::map<unsigned int,int> CMenuContainer::s_MenuScrolls;
CString CMenuContainer::s_MRUShortcuts[MRU_PROGRAMS_COUNT];
CComPtr<IShellFolder> CMenuContainer::s_pDesktop;
CComPtr<IKnownFolderManager> CMenuContainer::s_pKnownFolders;
HWND CMenuContainer::s_LastFGWindow;
HTHEME CMenuContainer::s_Theme;
HTHEME CMenuContainer::s_PagerTheme;
CWindow CMenuContainer::s_Tooltip;
CWindow CMenuContainer::s_TooltipBaloon;
int CMenuContainer::s_TipShowTime;
int CMenuContainer::s_TipHideTime;
int CMenuContainer::s_TipShowTimeFolder;
int CMenuContainer::s_TipHideTimeFolder;
int CMenuContainer::s_HotItem;
CMenuContainer *CMenuContainer::s_pHotMenu;
int CMenuContainer::s_TipItem;
CMenuContainer *CMenuContainer::s_pTipMenu;
RECT CMenuContainer::s_MainRect;
DWORD CMenuContainer::s_TaskbarState;
DWORD CMenuContainer::s_HoverTime;
DWORD CMenuContainer::s_XMouse;
DWORD CMenuContainer::s_SubmenuStyle;
CLIPFORMAT CMenuContainer::s_ShellFormat;
MenuSkin CMenuContainer::s_Skin;
std::vector<CMenuFader*> CMenuFader::s_Faders;

int CMenuContainer::CompareMenuString( const wchar_t *str1, const wchar_t *str2 )
{
	if (s_bLogicalSort)
		return StrCmpLogicalW(str1,str2);
	else
		return CompareString(LOCALE_USER_DEFAULT,LINGUISTIC_IGNORECASE,str1,-1,str2,-1)-CSTR_EQUAL;
}

CMenuContainer::CMenuContainer( CMenuContainer *pParent, int index, int options, const StdMenuItem *pStdItem, PIDLIST_ABSOLUTE path1, PIDLIST_ABSOLUTE path2, const CString &regName )
{
	m_RefCount=1;
	m_bSubMenu=(index>=0); // this may be true even if pParent is NULL (in case you want to show only sub-menus somewhere, use index=0 and pParent=NULL)
	m_HoverItem=-1;
	m_pParent=pParent;
	m_ParentIndex=pParent?index:-1;
	m_Options=options;
	m_pStdItem=pStdItem;
	m_RegName=regName;
	m_Bitmap=NULL;
	m_ArrowsBitmap[0]=m_ArrowsBitmap[1]=m_ArrowsBitmap[2]=m_ArrowsBitmap[3]=NULL;
	m_Region=NULL;
	m_Path1a[0]=ILCloneFull(path1);
	m_Path1a[1]=NULL;
	m_Path2a[0]=ILCloneFull(path2);
	m_Path2a[1]=NULL;

	if (options&CONTAINER_ALLPROGRAMS)
	{
		SHGetKnownFolderIDList(FOLDERID_Programs,0,NULL,&m_Path1a[1]);
		SHGetKnownFolderIDList(FOLDERID_CommonPrograms,0,NULL,&m_Path2a[1]);
	}

	m_rUser1.left=m_rUser1.right=0;
	m_rUser2.left=m_rUser2.right=0;

	ATLASSERT(path1 || !path2);
	if (!s_pDesktop)
		SHGetDesktopFolder(&s_pDesktop);

	if (!s_pKnownFolders)
		s_pKnownFolders.CoCreateInstance(CLSID_KnownFolderManager,NULL,CLSCTX_INPROC_SERVER);

	ATLASSERT(s_pDesktop);
	m_bDestroyed=false;
	s_Menus.push_back(this);
	m_Submenu=-1;
	m_bScrollTimer=false;

	CoCreateInstance(CLSID_DragDropHelper,NULL,CLSCTX_INPROC_SERVER,IID_IDropTargetHelper,(void**)&m_pDropTargetHelper);
	LOG_MENU(LOG_OPEN,L"Open Menu, ptr=%p, index=%d, options=%08X, name=%s",this,index,options,regName);
}

CMenuContainer::~CMenuContainer( void )
{
	for (std::vector<MenuItem>::iterator it=m_Items.begin();it!=m_Items.end();++it)
	{
		if (it->pItem1) ILFree(it->pItem1);
		if (it->pItem2) ILFree(it->pItem2);
	}
	if (m_Path1a[0]) ILFree(m_Path1a[0]);
	if (m_Path1a[1]) ILFree(m_Path1a[1]);
	if (m_Path2a[0]) ILFree(m_Path2a[0]);
	if (m_Path2a[1]) ILFree(m_Path2a[1]);
	if (std::find(s_Menus.begin(),s_Menus.end(),m_pParent)!=s_Menus.end()) // check if the parent is still alive
	{
		if (m_pParent->m_Submenu==m_ParentIndex)
		{
			if (!m_pParent->m_bDestroyed)
				m_pParent->InvalidateItem(m_ParentIndex);
			if (m_pParent->m_HotItem<0)
				m_pParent->SetHotItem(m_ParentIndex);
			m_pParent->m_Submenu=-1;
		}
	}
	if (m_Bitmap) DeleteObject(m_Bitmap);
	for (int i=0;i<_countof(m_ArrowsBitmap);i++)
		if (m_ArrowsBitmap[i]) DeleteObject(m_ArrowsBitmap[i]);
	if (m_Region) DeleteObject(m_Region);

	// must be here and not in OnDestroy because during drag/drop a menu can close while still processing messages
	s_Menus.erase(std::find(s_Menus.begin(),s_Menus.end(),this));
}

void CMenuContainer::AddFirstFolder( CComPtr<IShellFolder> pFolder, PIDLIST_ABSOLUTE path, std::vector<MenuItem> &items, int options, unsigned int hash0 )
{
	CComPtr<IEnumIDList> pEnum;
	if (!pFolder || pFolder->EnumObjects(NULL,SHCONTF_NONFOLDERS|SHCONTF_FOLDERS,&pEnum)!=S_OK) pEnum=NULL;

	PITEMID_CHILD pidl;
	while (pEnum && pEnum->Next(1,&pidl,NULL)==S_OK)
	{
		STRRET str;
		if (SUCCEEDED(pFolder->GetDisplayNameOf(pidl,SHGDN_INFOLDER|SHGDN_FORPARSING,&str)))
		{
			wchar_t *name;
			StrRetToStr(&str,pidl,&name);

			MenuItem item={MENU_NO};
			item.icon=g_IconManager.GetIcon(pFolder,pidl,(options&CONTAINER_LARGE)!=0);
			item.pItem1=ILCombine(path,pidl);
			bool bLibrary=_wcsicmp(PathFindExtension(name),L".library-ms")==0;

			if (SUCCEEDED(pFolder->GetDisplayNameOf(pidl,SHGDN_INFOLDER|SHGDN_NORMAL,&str)))
			{
				CharUpper(name);
				item.nameHash=CalcFNVHash(name,hash0);
				CoTaskMemFree(name);
				StrRetToStr(&str,pidl,&name);
				item.name=name;
				CoTaskMemFree(name);
			}
			else
			{
				item.name=name;
				CharUpper(name);
				item.nameHash=CalcFNVHash(name,hash0);
				CoTaskMemFree(name);
			}

			SFGAOF flags=SFGAO_FOLDER|SFGAO_STREAM|SFGAO_LINK; // check if the item is a folder, archive or a link
			if (FAILED(pFolder->GetAttributesOf(1,&pidl,&flags)))
				flags=0;
			if (bLibrary) flags&=~SFGAO_STREAM;
			item.bLink=(flags&SFGAO_LINK)!=0;
			item.bFolder=(!(options&CONTAINER_CONTROLPANEL) && !(options&CONTAINER_NOSUBFOLDERS) && (flags&SFGAO_FOLDER) && (!(flags&(SFGAO_STREAM|SFGAO_LINK)) || (s_bExpandLinks && item.bLink)));
			item.bPrograms=(options&CONTAINER_PROGRAMS)!=0;

			items.push_back(item);
#ifdef REPEAT_ITEMS
			for (int i=0;i<REPEAT_ITEMS;i++)
			{
				item.pItem1=ILCloneFull(item.pItem1);
				items.push_back(item);
			}
#endif
		}
		ILFree(pidl);
	}
}

void CMenuContainer::AddSecondFolder( CComPtr<IShellFolder> pFolder, PIDLIST_ABSOLUTE path, std::vector<MenuItem> &items, int options, unsigned int hash0 )
{
	CComPtr<IEnumIDList> pEnum;
	if (!pFolder || pFolder->EnumObjects(NULL,SHCONTF_NONFOLDERS|SHCONTF_FOLDERS,&pEnum)!=S_OK) pEnum=NULL;

	PITEMID_CHILD pidl;
	while (pEnum && pEnum->Next(1,&pidl,NULL)==S_OK)
	{
		STRRET str;
		if (SUCCEEDED(pFolder->GetDisplayNameOf(pidl,SHGDN_INFOLDER|SHGDN_FORPARSING,&str)))
		{
			wchar_t *name;
			StrRetToStr(&str,pidl,&name);
			CharUpper(name);
			unsigned int hash=CalcFNVHash(name,hash0);
			CoTaskMemFree(name);
			PIDLIST_ABSOLUTE pItem2=ILCombine(path,pidl);

			// look for another item with the same name
			bool bFound=false;
			for (std::vector<MenuItem>::iterator it=items.begin();it!=items.end();++it)
			{
				if (hash==it->nameHash)
				{
					it->pItem2=pItem2;
					bFound=true;
					break;
				}
			}

			if (!bFound)
			{
				bool bLibrary=_wcsicmp(PathFindExtension(name),L".library-ms")==0;

				STRRET str2;
				if (SUCCEEDED(pFolder->GetDisplayNameOf(pidl,SHGDN_INFOLDER|SHGDN_NORMAL,&str2)))
					StrRetToStr(&str2,pidl,&name);
				else
					StrRetToStr(&str,pidl,&name);

				// new item
				MenuItem item={MENU_NO};
				item.icon=g_IconManager.GetIcon(pFolder,pidl,(options&CONTAINER_LARGE)!=0);
				item.name=name;
				item.nameHash=hash;
				item.pItem1=pItem2;

				SFGAOF flags=SFGAO_FOLDER|SFGAO_STREAM|SFGAO_LINK;
				if (FAILED(pFolder->GetAttributesOf(1,&pidl,&flags)))
					flags=0;
				if (bLibrary) flags&=~SFGAO_STREAM;
				item.bLink=(flags&SFGAO_LINK)!=0;
				item.bFolder=(!(options&CONTAINER_CONTROLPANEL) && !(options&CONTAINER_NOSUBFOLDERS) && (flags&SFGAO_FOLDER) && (!(flags&(SFGAO_STREAM|SFGAO_LINK)) || (s_bExpandLinks && item.bLink)));
				item.bPrograms=(options&CONTAINER_PROGRAMS)!=0;

				items.push_back(item);
#ifdef REPEAT_ITEMS
				for (int i=0;i<REPEAT_ITEMS;i++)
				{
					item.pItem1=ILCloneFull(item.pItem1);
					items.push_back(item);
				}
#endif
				CoTaskMemFree(name);
			}
		}
		ILFree(pidl);
	}
}

// Initialize the m_Items list
void CMenuContainer::InitItems( void )
{
	m_Items.clear();
	m_bRefreshPosted=false;
	m_Submenu=-1;
	m_HotPos=GetMessagePos();
	m_ContextItem=-1;
	m_ScrollCount=0;

	if ((m_Options&CONTAINER_DOCUMENTS) && s_MaxRecentDocuments>0) // create the recent documents list
	{
		ATLASSERT(m_Path1a[0] && !m_Path2a[0]);

		// find all documents

		// with many recent files it takes a long time to go through the IShellFolder enumeration
		// so use FindFirstFile directly
		wchar_t recentPath[_MAX_PATH];
		SHGetPathFromIDList(m_Path1a[0],recentPath);
		wchar_t find[_MAX_PATH];
		Sprintf(find,_countof(find),L"%s\\*.lnk",recentPath);

		std::vector<Document> docs;

		WIN32_FIND_DATA data;
		HANDLE h=FindFirstFile(find,&data);
		while (h!=INVALID_HANDLE_VALUE)
		{
			Document doc;
			doc.name.Format(L"%s\\%s",recentPath,data.cFileName);
			doc.time=data.ftLastWriteTime;
			docs.push_back(doc);
			if (!FindNextFile(h,&data))
			{
				FindClose(h);
				break;
			}
		}

		// sort by time
		std::sort(docs.begin(),docs.end());
		if (m_Options&CONTAINER_SORTZA)
			std::reverse(docs.begin(),docs.end());

		CComPtr<IShellFolder> pFolder;
		s_pDesktop->BindToObject(m_Path1a[0],NULL,IID_IShellFolder,(void**)&pFolder);

		size_t count=0;
		CComPtr<IShellLink> pLink;
		if (SUCCEEDED(CoCreateInstance(CLSID_ShellLink,NULL,CLSCTX_INPROC_SERVER,IID_IShellLink,(LPVOID*)&pLink)))
		{
			CComQIPtr<IPersistFile> pFile=pLink;
			if (pFile)
			{
				// go through the items until we find s_MaxRecentDocuments documents
				for (std::vector<Document>::const_iterator it=docs.begin();it!=docs.end();++it)
				{
					wchar_t path[_MAX_PATH];
					// find the target of the lnk file
					if (SUCCEEDED(pFile->Load(it->name,STGM_READ)) && SUCCEEDED(pLink->GetPath(path,_countof(path),&data,0)))
					{
						// check if it is link to a file or directory
						if (path[0] && !(data.dwFileAttributes&FILE_ATTRIBUTE_DIRECTORY))
						{
							PIDLIST_ABSOLUTE pidl0=ILCreateFromPath(it->name);
							PITEMID_CHILD pidl=(PITEMID_CHILD)ILFindLastID(pidl0);
							STRRET str;
							if (SUCCEEDED(pFolder->GetDisplayNameOf(pidl,SHGDN_NORMAL,&str)))
							{
								wchar_t *name;
								StrRetToStr(&str,pidl,&name);

								MenuItem item={MENU_NO};
								item.icon=g_IconManager.GetIcon(pFolder,pidl,(m_Options&CONTAINER_LARGE)!=0);
								item.name=name;
								item.pItem1=pidl0;

								m_Items.push_back(item);
								CoTaskMemFree(name);
								count++;
							}
							else
								ILFree(pidl0);

							if ((int)count>=s_MaxRecentDocuments) break;
						}
					}
				}
			}
		}
	}

	// add first folder
	if (!(m_Options&CONTAINER_DOCUMENTS))
	{
		if (m_Path1a[0])
		{
			CComPtr<IShellFolder> pFolder;
			s_pDesktop->BindToObject(m_Path1a[0],NULL,IID_IShellFolder,(void**)&pFolder);
			m_pDropFoldera[0]=pFolder;
			AddFirstFolder(pFolder,m_Path1a[0],m_Items,m_Options,FNV_HASH0);
		}

		// add second folder
		if (m_Path2a[0])
		{
			CComPtr<IShellFolder> pFolder;
			s_pDesktop->BindToObject(m_Path2a[0],NULL,IID_IShellFolder,(void**)&pFolder);
			AddSecondFolder(pFolder,m_Path2a[0],m_Items,m_Options,FNV_HASH0);
		}
	}

	if (!m_pParent)
	{
		// remove the Programs subfolder from the main menu. it will be added separately
		PIDLIST_ABSOLUTE progs;
		if (FAILED(SHGetKnownFolderIDList(FOLDERID_Programs,0,NULL,&progs)))
			if (FAILED(SHGetKnownFolderIDList(FOLDERID_CommonPrograms,0,NULL,&progs)))
				progs=NULL;
		if (progs)
		{
			for (std::vector<MenuItem>::iterator it=m_Items.begin();it!=m_Items.end();++it)
				if (ILIsEqual(it->pItem1,progs))
				{
					ILFree(it->pItem1);
					if (it->pItem2) ILFree(it->pItem2);
					m_Items.erase(it);
					break;
				}
			ILFree(progs);
		}
	}

	if (m_Options&CONTAINER_ALLPROGRAMS)
	{
		std::vector<MenuItem> items;
		if (m_Path1a[1])
		{
			CComPtr<IShellFolder> pFolder;
			s_pDesktop->BindToObject(m_Path1a[1],NULL,IID_IShellFolder,(void**)&pFolder);
			m_pDropFoldera[1]=pFolder;
			AddFirstFolder(pFolder,m_Path1a[1],items,m_Options,CalcFNVHash(L"\\"));
		}

		// add second folder
		if (m_Path2a[1])
		{
			CComPtr<IShellFolder> pFolder;
			s_pDesktop->BindToObject(m_Path2a[1],NULL,IID_IShellFolder,(void**)&pFolder);
			AddSecondFolder(pFolder,m_Path2a[1],items,m_Options,CalcFNVHash(L"\\"));
		}
		if (!items.empty())
		{
			if (!m_Items.empty())
			{
				MenuItem item={MENU_SEPARATOR};
				item.priority=1;
				m_Items.push_back(item);
			}
			for (std::vector<MenuItem>::iterator it=items.begin();it!=items.end();++it)
				it->priority=2;
			m_Items.insert(m_Items.end(),items.begin(),items.end());
		}
	}

	// sort m_Items or read order from the registry
	LoadItemOrder();

	if (m_Items.size()>MAX_MENU_ITEMS)
	{
		for (size_t i=MAX_MENU_ITEMS;i<m_Items.size();i++)
		{
			if (m_Items[i].pItem1) ILFree(m_Items[i].pItem1);
			if (m_Items[i].pItem2) ILFree(m_Items[i].pItem2);
		}
		m_Items.resize(MAX_MENU_ITEMS);
	}

	if (m_Options&CONTAINER_CONTROLPANEL)
	{
		// expand Administrative Tools. must be done after the sorting because we don't want the folder to jump to the top
		unsigned int AdminToolsHash=CalcFNVHash(L"::{D20EA4E1-3957-11D2-A40B-0C5020524153}");
		for (std::vector<MenuItem>::iterator it=m_Items.begin();it!=m_Items.end();++it)
			if (it->nameHash==AdminToolsHash)
			{
				it->bFolder=true;
				break;
			}
	}

	if (m_Items.empty() && m_Path1a[0] && m_pDropFoldera[0])
	{
		// add (Empty) item to the empty submenus
		MenuItem item={m_bSubMenu?MENU_EMPTY:MENU_EMPTY_TOP};
		item.icon=I_IMAGENONE;
		item.name=FindTranslation(L"Menu.Empty",L"(Empty)");
		m_Items.push_back(item);
	}

	int recentType=GetSettingInt(L"RecentProgKeys");

	if (m_Options&CONTAINER_RECENT)
	{
		int nRecent=GetSettingInt(L"MaxRecentPrograms");
		bool bReverse=false;
		if (nRecent<0)
		{
			nRecent=-nRecent;
			bReverse=true;
		}
		if (nRecent>MRU_PROGRAMS_COUNT) nRecent=MRU_PROGRAMS_COUNT;
		if (nRecent>0)
		{
			LoadMRUShortcuts();
			// prepend recent programs
			std::vector<MenuItem> items;

			for (int i=0;i<MRU_PROGRAMS_COUNT;i++)
			{
				if (s_MRUShortcuts[i].IsEmpty()) break;
				PIDLIST_ABSOLUTE pItem;
				if (SUCCEEDED(s_pDesktop->ParseDisplayName(NULL,NULL,(wchar_t*)(const wchar_t*)s_MRUShortcuts[i],NULL,(PIDLIST_RELATIVE*)&pItem,NULL)))
				{
					CComPtr<IShellFolder> pFolder;
					PCUITEMID_CHILD pidl;
					if (SUCCEEDED(SHBindToParent(pItem,IID_IShellFolder,(void**)&pFolder,&pidl)))
					{
						STRRET str;
						if (SUCCEEDED(pFolder->GetDisplayNameOf(pidl,SHGDN_INFOLDER|SHGDN_NORMAL,&str)))
						{
							wchar_t *name;
							StrRetToStr(&str,pidl,&name);

							// new item
							MenuItem item={MENU_RECENT};
							item.icon=g_IconManager.GetIcon(pFolder,pidl,(m_Options&CONTAINER_LARGE)!=0);
							int idx=(int)items.size();
							if (idx<10)
							{
								if (recentType==2)
									item.name.Format(L"&%d %s",(idx+1)%10,name);
								else
									item.name=name;
								if (recentType==2 || recentType==3)
									item.accelerator=((idx+1)%10)+'0';
							}
							else
								item.name=name;
							item.nameHash=0;
							item.pItem1=pItem;
							pItem=NULL;

							items.push_back(item);
							CoTaskMemFree(name);
						}
					}
					if (pItem) ILFree(pItem);
				}
				if ((int)items.size()==nRecent)
					break;
			}

			if (!items.empty())
			{
				if (bReverse)
					std::reverse(items.begin(),items.end());
				MenuItem item={MENU_SEPARATOR};
				if (GetSettingBool(L"RecentProgsTop"))
				{
					items.push_back(item);
					m_Items.insert(m_Items.begin(),items.begin(),items.end());
				}
				else
				{
					m_Items.push_back(item);
					m_Items.insert(m_Items.end(),items.begin(),items.end());
				}
			}
		}
	}

	m_ScrollCount=(int)m_Items.size();

	// add standard items
	if (m_pStdItem && m_pStdItem->id!=MENU_NO)
	{
		if (!m_Items.empty())
		{
			MenuItem item={MENU_SEPARATOR};
			if (m_pStdItem->id==MENU_COLUMN_PADDING)
				item.bAlignBottom=true;
			if (m_Options&CONTAINER_ITEMS_FIRST)
				m_Items.insert(m_Items.begin(),item);
			else
				m_Items.push_back(item);
		}
		size_t menuIdx=(m_Options&CONTAINER_ITEMS_FIRST)?0:m_Items.size();
		bool bBreak=false, bAlignBottom=false;
		for (const StdMenuItem *pItem=m_pStdItem;pItem->id!=MENU_LAST;pItem++)
		{
			int stdOptions=MENU_ENABLED|MENU_EXPANDED;
			for (int i=0;i<_countof(g_StdOptions);i++)
				if (g_StdOptions[i].id==pItem->id)
				{
					stdOptions=g_StdOptions[i].options;
					break;
				}

			if (!(stdOptions&MENU_ENABLED)) continue;

			if (pItem->id==MENU_COLUMN_BREAK)
			{
				bBreak=true;
				continue;
			}
			if (pItem->id==MENU_COLUMN_PADDING)
			{
				bAlignBottom=true;
				continue;
			}

			MenuItem item={pItem->id,pItem};

			item.bBreak=bBreak;
			item.bAlignBottom=bAlignBottom;
			bBreak=bAlignBottom=false;

			ATLASSERT(pItem->folder1 || !pItem->folder2);
			if (pItem->folder1)
			{
				SHGetKnownFolderIDList(*pItem->folder1,0,NULL,&item.pItem1);
				if (pItem->folder2)
					SHGetKnownFolderIDList(*pItem->folder2,0,NULL,&item.pItem2);
				wchar_t recentPath[_MAX_PATH];
				SHGetPathFromIDList(item.pItem1,recentPath);
				item.bFolder=(stdOptions&MENU_EXPANDED)!=0;
			}
			else if (pItem->link)
			{
				SFGAOF flags=SFGAO_FOLDER|SFGAO_STREAM|SFGAO_LINK;
				wchar_t buf[1024];
				Strcpy(buf,_countof(buf),item.pStdItem->link);
				DoEnvironmentSubst(buf,_countof(buf));
				bool bLibrary=_wcsicmp(PathFindExtension(buf),L".library-ms")==0;
				if (SUCCEEDED(s_pDesktop->ParseDisplayName(NULL,NULL,buf,NULL,(PIDLIST_RELATIVE*)&item.pItem1,&flags)))
				{
					if (bLibrary) flags&=~SFGAO_STREAM;
					item.bLink=(flags&SFGAO_LINK)!=0;
					item.bFolder=((flags&SFGAO_FOLDER) && !(item.pStdItem->settings&StdMenuItem::MENU_NOEXPAND) && (!(flags&(SFGAO_STREAM|SFGAO_LINK)) || (s_bExpandLinks && item.bLink)));
				}
			}
			if ((pItem->submenu && (stdOptions&MENU_EXPANDED)) || pItem->id==MENU_RECENT_ITEMS)
				item.bFolder=true;

			// get icon
			if (pItem->iconPath)
			{
				if (_wcsicmp(pItem->iconPath,L"none")==0)
					item.icon=I_IMAGENONE; 
				else
					item.icon=g_IconManager.GetCustomIcon(pItem->iconPath,(m_Options&CONTAINER_LARGE)!=0);
			}
			else if (item.pItem1)
			{
				// for some reason SHGetFileInfo(SHGFI_PIDL|SHGFI_ICONLOCATION) doesn't work here. go to the IShellFolder directly
				CComPtr<IShellFolder> pFolder2;
				PCUITEMID_CHILD pidl;
				if (SUCCEEDED(SHBindToParent(item.pItem1,IID_IShellFolder,(void**)&pFolder2,&pidl)))
					item.icon=g_IconManager.GetIcon(pFolder2,pidl,(m_Options&CONTAINER_LARGE)!=0);
			}

			// get name
			if (pItem->label)
			{
				if (item.id==MENU_LOGOFF)
				{
					// construct the text Log Off <username>...
					wchar_t user[256]={0};
					ULONG size=_countof(user);
					if (!GetUserNameEx(NameDisplay,user,&size))
					{
						// GetUserNameEx may fail (for example on Home editions). use the login name
						size=_countof(user);
						GetUserName(user,&size);
					}
					item.name.Format(pItem->label,user);
				}
				else
					item.name=pItem->label;
			}
			else if (item.pItem1)
			{
				SHFILEINFO info={0};
				SHGetFileInfo((LPCWSTR)item.pItem1,0,&info,sizeof(info),SHGFI_PIDL|SHGFI_DISPLAYNAME);
				item.name=info.szDisplayName;
			}

			item.bPrograms=(item.id==MENU_PROGRAMS || item.id==MENU_FAVORITES);
			m_Items.insert(m_Items.begin()+menuIdx,1,item);
			menuIdx++;
		}
	}

	if (m_Items.empty())
	{
		// add (Empty) item to the empty submenus
		MenuItem item={MENU_EMPTY};
		item.icon=I_IMAGENONE;
		item.name=FindTranslation(L"Menu.Empty",L"(Empty)");
		m_Items.push_back(item);
	}

	// remove trailing separators
	while (!m_Items.empty() && m_Items[m_Items.size()-1].id==MENU_SEPARATOR)
		m_Items.pop_back();

	if (m_bSubMenu)
		m_ScrollCount=(int)m_Items.size();

	// find accelerators
	for (std::vector<MenuItem>::iterator it=m_Items.begin();it!=m_Items.end();++it)
	{
		if (it->id==MENU_SEPARATOR || it->id==MENU_EMPTY  || it->id==MENU_EMPTY_TOP || it->name.IsEmpty() || (it->id==MENU_RECENT && recentType!=1))
			continue;

		const wchar_t *name=it->name;
		wchar_t buf[2]={name[0],0};
		while (1)
		{
			const wchar_t *c=wcschr(name,'&');
			if (!c || !c[1]) break;
			if (c[1]!='&')
			{
				buf[0]=c[1];
				break;
			}
			name=c+1;
		}
		CharUpper(buf); // always upper case
		it->accelerator=buf[0];
	}
}

// Calculate the size and create the background bitmaps
void CMenuContainer::InitWindow( void )
{
	m_bTwoColumns=(!m_bSubMenu && s_Skin.bTwoColumns);
	if (!m_pParent && !s_Theme && IsAppThemed())
	{
		s_Theme=OpenThemeData(m_hWnd,L"toolbar");
		s_PagerTheme=OpenThemeData(m_hWnd,L"scrollbar");
	}
	if (!m_pParent && !s_Tooltip.m_hWnd)
	{
		s_Tooltip=CreateWindowEx(WS_EX_TOPMOST|WS_EX_TOOLWINDOW|WS_EX_TRANSPARENT|(s_bRTL?WS_EX_LAYOUTRTL:0),TOOLTIPS_CLASS,NULL,WS_POPUP|TTS_NOPREFIX|TTS_ALWAYSTIP,0,0,0,0,NULL,NULL,g_Instance,NULL);
		s_Tooltip.SendMessage(TTM_SETMAXTIPWIDTH,0,500);
		TOOLINFO tool={sizeof(tool),TTF_ABSOLUTE|TTF_TRACK|TTF_TRANSPARENT|(s_bRTL?TTF_RTLREADING:0)};
		tool.uId=1;
		s_Tooltip.SendMessage(TTM_ADDTOOL,0,(LPARAM)&tool);
	}

	{
		unsigned colors[4]={m_bSubMenu?s_Skin.Submenu_arrow_color[0]:s_Skin.Main_arrow_color[0],m_bSubMenu?s_Skin.Submenu_arrow_color[1]:s_Skin.Main_arrow_color[1],s_Skin.Main_arrow_color2[0],s_Skin.Main_arrow_color2[1]};
		for (int i=0;i<4;i++)
		{
			if (!m_ArrowsBitmap[i])
			{
				unsigned int color=colors[i];
				color=0xFF000000|(color<<16)|(color&0xFF00)|((color>>16)&0xFF);
				m_ArrowsBitmap[i]=(HBITMAP)LoadImage(g_Instance,MAKEINTRESOURCE(IDB_ARROWS),IMAGE_BITMAP,0,0,LR_CREATEDIBSECTION);
				BITMAP info;
				GetObject(m_ArrowsBitmap[i],sizeof(info),&info);
				int n=info.bmWidth*info.bmHeight;
				for (int p=0;p<n;p++)
					((unsigned int*)info.bmBits)[p]&=color;
			}
		}
	}
	// calculate maximum height
	int maxHeight;
	int maxWidth=m_MaxWidth;
	int mainBottom=s_MainRect.bottom;
	int mainTop=s_MainRect.top;
	{
		maxHeight=(mainBottom-mainTop);
		// adjust for padding
		RECT rc={0,0,0,0};
		AdjustWindowRect(&rc,GetWindowLong(GWL_STYLE),FALSE);
		maxWidth-=rc.right-rc.left;
		maxWidth-=s_Skin.Submenu_padding.left+s_Skin.Submenu_padding.right;
		if (m_bSubMenu)
		{
			mainTop-=s_Skin.Submenu_padding.top-rc.top;
			mainBottom+=rc.bottom+s_Skin.Submenu_padding.bottom;
		}
		else
		{
			RECT rc2;
			GetWindowRect(&rc2);
			if (m_Options&CONTAINER_TOP)
				maxHeight=mainBottom-rc2.top;
			else
				maxHeight=rc2.bottom-mainTop;
			maxHeight-=rc.bottom-rc.top;
			maxHeight-=s_Skin.Main_padding.top+s_Skin.Main_padding.bottom;
		}
	}
#ifdef _DEBUG
//	maxHeight/=3; // uncomment to test for smaller screen
#endif

	int iconSize=(m_Options&CONTAINER_LARGE)?CIconManager::LARGE_ICON_SIZE:CIconManager::SMALL_ICON_SIZE;

	HDC hdc=CreateCompatibleDC(NULL);
	m_Font[0]=m_bSubMenu?s_Skin.Submenu_font:s_Skin.Main_font;
	m_Font[1]=s_Skin.Main_font2;
	const RECT iconPadding[2]={m_bSubMenu?s_Skin.Submenu_icon_padding:s_Skin.Main_icon_padding,s_Skin.Main_icon_padding2};
	const RECT textPadding[2]={m_bSubMenu?s_Skin.Submenu_text_padding:s_Skin.Main_text_padding,s_Skin.Main_text_padding2};
	int arrowSize[2];
	if (m_bSubMenu)
		arrowSize[0]=s_Skin.Submenu_arrow_padding.cx+s_Skin.Submenu_arrow_padding.cy+s_Skin.Submenu_arrow_Size.cx;
	else
		arrowSize[0]=s_Skin.Main_arrow_padding.cx+s_Skin.Main_arrow_padding.cy+s_Skin.Main_arrow_Size.cx;
	arrowSize[1]=s_Skin.Main_arrow_padding2.cx+s_Skin.Main_arrow_padding2.cy+s_Skin.Main_arrow_Size2.cx;

	HFONT font0=(HFONT)SelectObject(hdc,m_Font[m_bTwoColumns?1:0]);
	TEXTMETRIC metrics;
	GetTextMetrics(hdc,&metrics);
	if (m_bTwoColumns)
	{
		int textHeight=metrics.tmHeight+textPadding[1].top+textPadding[1].bottom;
		int itemHeight=iconSize+iconPadding[1].top+iconPadding[1].bottom;
		if (itemHeight<textHeight)
		{
			m_IconTopOffset[1]=(textHeight-itemHeight)/2;
			m_TextTopOffset[1]=0;
			itemHeight=textHeight;
		}
		else
		{
			m_IconTopOffset[1]=0;
			m_TextTopOffset[1]=(itemHeight-textHeight)/2;
		}
		m_ItemHeight[1]=itemHeight;

		int numChar=GetSettingInt(L"MaxMainMenuWidth");
		m_MaxItemWidth[1]=numChar?metrics.tmAveCharWidth*numChar:65536;
		SelectObject(hdc,m_Font[0]);
		GetTextMetrics(hdc,&metrics);
	}
	{
		int textHeight=metrics.tmHeight+textPadding[0].top+textPadding[0].bottom;
		int itemHeight=iconSize+iconPadding[0].top+iconPadding[0].bottom;
		if (itemHeight<textHeight)
		{
			m_IconTopOffset[0]=(textHeight-itemHeight)/2;
			m_TextTopOffset[0]=0;
			itemHeight=textHeight;
		}
		else
		{
			m_IconTopOffset[0]=0;
			m_TextTopOffset[0]=(itemHeight-textHeight)/2;
		}
		m_ItemHeight[0]=itemHeight;

		int numChar=GetSettingInt(m_bSubMenu?L"MaxMenuWidth":L"MaxMainMenuWidth");
		m_MaxItemWidth[0]=numChar?metrics.tmAveCharWidth*numChar:65536;
	}
	m_ScrollButtonSize=m_ItemHeight[0]/2;
	if (m_ScrollButtonSize<MIN_SCROLL_HEIGHT) m_ScrollButtonSize=MIN_SCROLL_HEIGHT;

	int sepHeight[2]={SEPARATOR_HEIGHT,SEPARATOR_HEIGHT};
	
	if (m_bSubMenu)
	{
		if (s_Skin.Submenu_separator) sepHeight[0]=s_Skin.Submenu_separatorHeight;
	}
	else
	{
		if (s_Skin.Main_separator) sepHeight[0]=s_Skin.Main_separatorHeight;
		if (s_Skin.Main_separator2) sepHeight[1]=s_Skin.Main_separatorHeight2;
	}

	// calculate item size
	std::vector<int> columnWidths;
	columnWidths.push_back(0);

	bool bMultiColumn=s_ScrollMenus!=0 && (m_Options&CONTAINER_MULTICOLUMN);

	{
		int row=0, column=0;
		int y=0;
		int maxw=0;
		int index=0;
		for (size_t i=0;i<m_Items.size();i++)
		{
			MenuItem &item=m_Items[i];
			if (m_bTwoColumns && column==0 && i>0 && m_Items[i].bBreak)
			{
				// start a new column
				column++;
				columnWidths.push_back(0);
				row=0;
				y=0;
				index=1;
				SelectObject(hdc,m_Font[1]);
			}
			int w=0, h=0;
			if (!s_bShowTopEmpty && m_Items.size()>1 && (m_Items[i].id==MENU_EMPTY_TOP || (i>0 && m_Items[i-1].id==MENU_EMPTY_TOP)))
			{
				h=0; // this is the first (Empty) item in the top menu. hide it for now
			}
			else if (item.id==MENU_SEPARATOR)
			{
				if (y==0)
					h=0; // ignore separators at the top of the column
				else
					h=sepHeight[index];
			}
			else
			{
				h=m_ItemHeight[index];
				SIZE size;
				if (GetTextExtentPoint32(hdc,item.name,item.name.GetLength(),&size))
					w=size.cx;
				if (w>m_MaxItemWidth[index]) w=m_MaxItemWidth[index];
				w+=iconSize+iconPadding[index].left+iconPadding[index].right+textPadding[index].left+textPadding[index].right+arrowSize[index];
			}
			if (bMultiColumn && y>0 && y+h>maxHeight)
			{
				if (item.id==MENU_SEPARATOR)
					h=0; // ignore separators at the bottom of the column
				else
				{
					// start a new column
					column++;
					columnWidths.push_back(0);
					row=0;
					y=0;
				}
			}
			else if (item.id==MENU_SEPARATOR && m_bTwoColumns && column==0 && i+1<m_Items.size() && m_Items[i+1].bBreak)
				h=0;
			if (columnWidths[column]<w) columnWidths[column]=w;
			item.row=row;
			item.column=column;
			item.itemRect.top=y;
			item.itemRect.bottom=y+h;
			item.itemRect.left=0;
			item.itemRect.right=w;
			y+=h;
			row++;
		}
	}

	SelectObject(hdc,font0);
	DeleteDC(hdc);

	if (columnWidths.size()==1)
		m_bTwoColumns=false;

	// calculate width of each column
	if (!m_bTwoColumns && GetSettingBool(L"SameSizeColumns"))
	{
		int maxw=0;
		for (size_t i=0;i<columnWidths.size();i++)
			if (maxw<columnWidths[i])
				maxw=columnWidths[i];
		for (size_t i=0;i<columnWidths.size();i++)
			columnWidths[i]=maxw;
	}

	if (s_ScrollMenus==2 && columnWidths.size()>1 && m_bSubMenu)
	{
		// auto - determine if we should have 1 column or many
		int width=0;
		for (size_t i=0;i<columnWidths.size();i++)
		{
			if (i>0) width+=s_Skin.Submenu_separatorWidth;
			width+=columnWidths[i];
		}
		if (width>maxWidth)
		{
			bMultiColumn=false;
			// the columns don't fit on screen, switch to one scrollable column
			int y=0;
			columnWidths.resize(1);
			columnWidths[0]=0;
			for (size_t i=0;i<m_Items.size();i++)
			{
				MenuItem &item=m_Items[i];
				int h=0;
				if (item.id==MENU_SEPARATOR)
				{
					if (y==0)
						h=0; // ignore separators at the top of the column
					else
						h=sepHeight[0];
				}
				else
				{
					h=m_ItemHeight[0];
				}
				if (columnWidths[0]<item.itemRect.right) columnWidths[0]=item.itemRect.right;
				item.row=(int)i;
				item.column=0;
				item.itemRect.top=y;
				item.itemRect.bottom=y+h;
				y+=h;
			}
		}
	}

	// calculate the horizontal position of each item
	int maxw=0;
	int maxh=0;
	{
		m_ColumnOffsets.resize(columnWidths.size());
		for (size_t i=0;i<columnWidths.size();i++)
		{
			if (i>0) maxw+=s_Skin.Submenu_separatorWidth;
			m_ColumnOffsets[i]=maxw;
			maxw+=columnWidths[i];
		}
		columnWidths.push_back(maxw);
		int x=0;
		for (size_t i=0;i<m_Items.size();i++)
		{
			MenuItem &item=m_Items[i];
			item.itemRect.left=m_ColumnOffsets[item.column];
			item.itemRect.right=item.itemRect.left+columnWidths[item.column];
			if (maxh<item.itemRect.bottom) maxh=item.itemRect.bottom;
		}
	}

	if (m_Bitmap)
	{
		DeleteObject(m_Bitmap);
		m_Bitmap=NULL;
	}
	if (m_Region)
	{
		DeleteObject(m_Region);
		m_Region=NULL;
	}

	int totalWidth, totalHeight;
	memset(&m_rContent2,0,sizeof(m_rContent2));
	{
		int w1=maxw, w2=0;
		int h1=(maxh<maxHeight?maxh:maxHeight), h2=0;
		if (m_bTwoColumns && columnWidths.size()>2)
		{
			w1=columnWidths[0];
			w2=columnWidths[1];
			h2=m_Items[m_Items.size()-1].itemRect.bottom;
		}
		if (!m_bSubMenu)
		{
			if (s_Skin.Main_bitmap || s_Skin.User_image_size || m_bTwoColumns || s_Skin.User_name_position.left!=s_Skin.User_name_position.right)
			{
				CreateBackground(w1,w2,h1,h2);
				BITMAP info;
				GetObject(m_Bitmap,sizeof(info),&info);
				totalWidth=info.bmWidth;
				totalHeight=info.bmHeight;
			}
			else
			{
				m_rContent.left=s_Skin.Main_padding.left;
				m_rContent.top=s_Skin.Main_padding.top;
				m_rContent.right=s_Skin.Main_padding.left+w1;
				m_rContent.bottom=s_Skin.Main_padding.top+h1;
				totalWidth=s_Skin.Main_padding.left+s_Skin.Main_padding.right+w1;
				totalHeight=s_Skin.Main_padding.top+s_Skin.Main_padding.bottom+h1;
			}
		}
		else
		{
			if (s_Skin.Submenu_bitmap)
				CreateSubmenuRegion(w1,h1);

			m_rContent.left=s_Skin.Submenu_padding.left;
			m_rContent.top=s_Skin.Submenu_padding.top;
			m_rContent.right=s_Skin.Submenu_padding.left+w1;
			m_rContent.bottom=s_Skin.Submenu_padding.top+h1;
			totalWidth=s_Skin.Submenu_padding.left+s_Skin.Submenu_padding.right+w1;
			totalHeight=s_Skin.Submenu_padding.top+s_Skin.Submenu_padding.bottom+h1;
		}
		// offset the items
		for (size_t i=0;i<m_Items.size();i++)
		{
			int dx=m_rContent.left;
			int dy=m_rContent.top;
			if (m_bTwoColumns && m_Items[i].column==1)
			{
				dx=m_rContent2.left-m_Items[i].itemRect.left;
				dy=m_rContent2.top;
			}
			OffsetRect(&m_Items[i].itemRect,dx,dy);
		}

		if (m_bTwoColumns && columnWidths.size()>2)
		{
			int dh1=0, dh2=0;
			for (size_t i=0;i<m_Items.size();i++)
			{
				if (m_Items[i].column==0)
					dh1=m_Items[i].itemRect.bottom;
				else
					dh2=m_Items[i].itemRect.bottom;
			}
			m_rContent.bottom=totalHeight-s_Skin.Main_padding.bottom;
			m_rContent2.bottom=totalHeight-s_Skin.Main_padding2.bottom;
			dh1=m_rContent.bottom-dh1;
			if (dh1<0) dh1=0;
			dh2=m_rContent2.bottom-dh2;
			if (dh2<0) dh2=0;

			bool bAlign1=false, bAlign2=false;
			for (size_t i=0;i<m_Items.size();i++)
			{
				if (m_Items[i].column==0)
				{
					if (m_Items[i].bAlignBottom)
						bAlign1=true;
					if (bAlign1)
						OffsetRect(&m_Items[i].itemRect,0,dh1);
				}
				else
				{
					if (m_Items[i].bAlignBottom)
						bAlign2=true;
					if (bAlign2)
						OffsetRect(&m_Items[i].itemRect,0,dh2);
				}
			}
		}
	}

	// create pager
	if (!bMultiColumn && maxh>maxHeight && m_ScrollCount>0)
	{
		int d=maxh-maxHeight;
		for (size_t i=m_ScrollCount;i<m_Items.size();i++)
			if (m_Items[i].column==0)
				OffsetRect(&m_Items[i].itemRect,0,-d);

		std::map<unsigned int,int>::iterator it=s_MenuScrolls.find(CalcFNVHash(m_RegName));
		if (it!=s_MenuScrolls.end())
		{
			m_ScrollOffset=it->second; // restore the scroll position if the same menu has been opened before
			if (m_ScrollOffset>d) m_ScrollOffset=d;
		}
		else
			m_ScrollOffset=0;
		m_ScrollHeight=m_Items[m_ScrollCount-1].itemRect.bottom-d-m_rContent.top;
	}
	else
		m_ScrollOffset=m_ScrollHeight=0;
	UpdateScroll();
	m_bScrollUpHot=m_bScrollDownHot=false;

	RECT rc={0,0,totalWidth,totalHeight};
	AdjustWindowRect(&rc,GetWindowLong(GWL_STYLE),FALSE);
	RECT rc0;
	GetWindowRect(&rc0);
	OffsetRect(&rc,(m_Options&CONTAINER_LEFT)?(rc0.left-rc.left):(rc0.right-rc.right),(m_Options&CONTAINER_TOP)?(rc0.top-rc.top):(rc0.bottom-rc.bottom));
	if (m_bSubMenu)
	{
		int dx=0,dy=0;
		if (rc.bottom>mainBottom)
			dy=mainBottom-rc.bottom;
		if (rc.top+dy<mainTop)
			dy=mainTop-rc.top;
		OffsetRect(&rc,dx,dy);
	}
	SetWindowPos(NULL,&rc,SWP_NOZORDER|SWP_NOACTIVATE|SWP_NOCOPYBITS|SWP_DEFERERASE);

	// for some reason the region must be set after the call to SetWindowPos. otherwise it doesn't work for RTL windows
	MenuSkin::TOpacity opacity=(m_bSubMenu?s_Skin.Submenu_opacity:s_Skin.Main_opacity);
	if (m_Region && (opacity==MenuSkin::OPACITY_REGION || opacity==MenuSkin::OPACITY_GLASS || opacity==MenuSkin::OPACITY_FULLGLASS))
	{
		int size=GetRegionData(m_Region,0,NULL);
		std::vector<char> buf(size);
		GetRegionData(m_Region,size,(RGNDATA*)&buf[0]);
		XFORM xform={1,0,0,1};
		if (s_bRTL)
		{
			// mirror the region (again)
			xform.eM11=-1;
			xform.eDx=(float)(rc.right-rc.left);
		}
		HRGN rgn=ExtCreateRegion(&xform,size,(RGNDATA*)&buf[0]);

		if (!SetWindowRgn(rgn))
			DeleteObject(rgn); // otherwise the OS takes ownership of the region, no need to free
	}

	m_bTrackMouse=false;
	m_bScrollTimer=false;
	m_InsertMark=-1;
	m_HotItem=-1;
	m_Submenu=-1;
	m_MouseWheel=0;

	if (!m_bSubMenu)
	{
		TOOLINFO tool={sizeof(tool),TTF_SUBCLASS|TTF_TRANSPARENT|(s_bRTL?TTF_RTLREADING:0)};
		tool.hwnd=m_hWnd;
		tool.uId=2;
		if (m_rUser1.left<m_rUser1.right || m_rUser2.left<m_rUser2.right)
		{
			// construct the text Log Off <username>...
			wchar_t user[256]={0};
			ULONG size=_countof(user);
			if (!GetUserNameEx(NameDisplay,user,&size))
			{
				// GetUserNameEx may fail (for example on Home editions). use the login name
				DWORD size=_countof(user);
				GetUserName(user,&size);
			}
			tool.lpszText=user;

			if (m_rUser1.left<m_rUser1.right)
			{
				tool.rect=m_rUser1;
				s_Tooltip.SendMessage(TTM_ADDTOOL,0,(LPARAM)&tool);
			}
			else
				s_Tooltip.SendMessage(TTM_DELTOOL,0,(LPARAM)&tool);

			tool.uId=3;
			if (m_rUser2.left<m_rUser2.right)
			{
				tool.rect=m_rUser2;
				s_Tooltip.SendMessage(TTM_ADDTOOL,0,(LPARAM)&tool);
			}
			else
				s_Tooltip.SendMessage(TTM_DELTOOL,0,(LPARAM)&tool);
		}
		else
		{
			s_Tooltip.SendMessage(TTM_DELTOOL,0,(LPARAM)&tool);
			tool.uId=3;
			s_Tooltip.SendMessage(TTM_DELTOOL,0,(LPARAM)&tool);
		}
	}
}

void CMenuContainer::UpdateScroll( void )
{
	if (m_ScrollHeight==0)
		m_bScrollUp=m_bScrollDown=false;
	else
	{
		m_bScrollUp=(m_ScrollOffset>0);
		m_bScrollDown=(m_ScrollOffset+m_ScrollHeight<m_Items[m_ScrollCount-1].itemRect.bottom-m_rContent.top);
	}
}

void CMenuContainer::UpdateScroll( const POINT *pt )
{
	if (m_bScrollUp)
	{
		RECT rc=m_rContent;
		rc.bottom=rc.top+m_ScrollButtonSize;
		bool bHot=pt && PtInRect(&rc,*pt);
		if (m_bScrollUpHot!=bHot)
		{
			m_bScrollUpHot=bHot;
			Invalidate();
		}
	}
	else
		m_bScrollUpHot=false;

	if (m_bScrollDown)
	{
		RECT rc=m_rContent;
		rc.bottom=m_rContent.top+m_ScrollHeight;
		rc.top=rc.bottom-m_ScrollButtonSize;
		bool bHot=pt && PtInRect(&rc,*pt);
		if (m_bScrollDownHot!=bHot)
		{
			m_bScrollDownHot=bHot;
			Invalidate();
		}
	}
	else
		m_bScrollDownHot=false;

	if (m_bScrollUpHot || m_bScrollDownHot)
	{
		if (!m_bScrollTimer)
		{
			m_bScrollTimer=true;
			SetTimer(TIMER_SCROLL,50);
		}
	}
	else if (m_bScrollTimer)
	{
		m_bScrollTimer=false;
		KillTimer(TIMER_SCROLL);
	}
}

LRESULT CMenuContainer::OnCreate( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled )
{
	MenuSkin::TOpacity opacity=(m_bSubMenu?s_Skin.Submenu_opacity:s_Skin.Main_opacity);
	if (opacity==MenuSkin::OPACITY_ALPHA || opacity==MenuSkin::OPACITY_FULLALPHA)
	{
		MARGINS margins={-1};
		DwmExtendFrameIntoClientArea(m_hWnd,&margins);
	}
	if (!m_pParent)
		BufferedPaintInit();
	InitWindow();
	m_HotPos=GetMessagePos();
	if (GetSettingBool(L"EnableAccessibility"))
	{
		m_pAccessible=new CMenuAccessible(this);
		NotifyWinEvent(EVENT_SYSTEM_MENUPOPUPSTART,m_hWnd,OBJID_CLIENT,CHILDID_SELF);
	}
	else
		m_pAccessible=NULL;
	RegisterDragDrop(m_hWnd,this);
	PlayMenuSound(m_bSubMenu?SOUND_POPUP:SOUND_MAIN);
	return 0;
}

bool CMenuContainer::GetItemRect( int index, RECT &rc )
{
	if (index>=0 && index<=(int)m_Items.size())
	{
		rc=m_Items[index].itemRect;
		if (m_ScrollHeight>0 && index<m_ScrollCount)
		{
			OffsetRect(&rc,0,-m_ScrollOffset);
			if (m_bScrollUp && rc.bottom<=m_rContent.top+m_ScrollButtonSize)
				return false;
			if (m_bScrollDown && rc.top>=m_rContent.top+m_ScrollHeight-m_ScrollButtonSize)
				return false;
		}
	}
	return true;
}

int CMenuContainer::HitTest( const POINT &pt, bool bDrop )
{
	if (m_bScrollUp && pt.y<m_rContent.top+m_ScrollButtonSize)
		return -1;
	int start=0;
	if (m_bScrollDown && pt.y>=m_rContent.top+m_ScrollHeight-m_ScrollButtonSize)
		start=m_ScrollCount;
	int n=(int)m_Items.size();
	for (int i=start;i<n;i++)
	{
		RECT rc=m_Items[i].itemRect;
		if (m_ScrollHeight>0 && i<m_ScrollCount)
		{
			OffsetRect(&rc,0,-m_ScrollOffset);
		}
		else if (bDrop && m_bTwoColumns && i<n-1 && m_Items[i+1].column==0 && m_Items[i+1].bAlignBottom)
			rc.bottom=m_Items[i+1].itemRect.top; // when dropping on the padding of the first column, assume the item above the padding was hit
		if (PtInRect(&rc,pt))
			return i;
	}
	return -1;
}

void CMenuContainer::InvalidateItem( int index )
{
	if (index>=0)
	{
		RECT rc;
		GetItemRect(index,rc);
		InvalidateRect(&rc);
	}
}

void CMenuContainer::SetHotItem( int index )
{
	if (index==m_HotItem) return;
	if ((index>=0)!=(m_HotItem>=0))
	{
		InvalidateItem(m_Submenu);
		InvalidateItem(m_ContextItem);
	}
	InvalidateItem(m_HotItem);
	InvalidateItem(index);
	m_HotItem=index;
	s_pTipMenu=NULL;
	s_TipItem=-1;
	if (index>=0)
	{
		s_pHotMenu=this;
		s_HotItem=index;
	}
	else if (s_pHotMenu==this)
	{
		s_pHotMenu=NULL;
		s_HotItem=-1;
		if (s_Tooltip.m_hWnd)
		{
			TOOLINFO tool={sizeof(tool),TTF_ABSOLUTE|TTF_TRACK|TTF_TRANSPARENT};
			tool.uId=1;
			s_Tooltip.SendMessage(TTM_TRACKACTIVATE,FALSE,(LPARAM)&tool);
		}
	}
	else
		return;
	if (index>=0 && index<(int)m_Items.size())
	{
		int show, hide;
		if (m_Items[index].bFolder)
			show=s_TipShowTimeFolder, hide=s_TipHideTimeFolder;
		else
			show=s_TipShowTime, hide=s_TipHideTime;
		if (s_Tooltip.m_hWnd)
		{
			TOOLINFO tool={sizeof(tool),TTF_ABSOLUTE|TTF_TRACK|TTF_TRANSPARENT};
			tool.uId=1;
			s_Tooltip.SendMessage(TTM_TRACKACTIVATE,FALSE,(LPARAM)&tool);
			if (!s_Menus[0]->m_bDestroyed && hide>0)
			{
				s_pTipMenu=s_pHotMenu;
				s_TipItem=s_HotItem;
				s_Menus[0]->SetTimer(TIMER_TOOLTIP_SHOW,show);
			}
		}
		NotifyWinEvent(EVENT_OBJECT_FOCUS,m_hWnd,OBJID_CLIENT,index+1);
	}
}

void CMenuContainer::SetInsertMark( int index, bool bAfter )
{
	if (index==m_InsertMark && bAfter==m_bInsertAfter) return;
	RECT rc;
	if (GetInsertRect(rc))
		InvalidateRect(&rc);
	m_InsertMark=index;
	m_bInsertAfter=bAfter;
	if (GetInsertRect(rc))
		InvalidateRect(&rc);
}

bool CMenuContainer::GetInsertRect( RECT &rc )
{
	if (m_InsertMark<0 || m_InsertMark>=(int)m_Items.size())
		return false;
	const MenuItem &item=m_Items[m_InsertMark];
	rc=item.itemRect;
	if (m_bInsertAfter)
		rc.top=rc.bottom;
	if (m_ScrollHeight>0 && m_InsertMark<m_ScrollCount)
		rc.top-=m_ScrollOffset;
	rc.top-=3;
	rc.bottom=rc.top+6;
	return true;
}

LRESULT CMenuContainer::OnSetContextItem( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled )
{
	for (int i=0;i<(int)m_Items.size();i++)
	{
		if (m_Items[i].nameHash==wParam)
		{
			m_ContextItem=i;
			break;
		}
	}
	return 0;
}

// Turn on the keyboard cues from now on. This is done when a keyboard action is detected
void CMenuContainer::ShowKeyboardCues( void )
{
	if (!s_bKeyboardCues)
	{
		s_bKeyboardCues=true;
		for (std::vector<CMenuContainer*>::const_iterator it=s_Menus.begin();it!=s_Menus.end();++it)
			(*it)->Invalidate();
	}
}

void CMenuContainer::SetActiveWindow( void )
{
	if (GetActiveWindow()!=m_hWnd)
		::SetActiveWindow(m_hWnd);
	if (!m_bSubMenu && s_bBehindTaskbar)
		SetWindowPos(g_TaskBar,0,0,0,0,SWP_NOSIZE|SWP_NOMOVE); // make sure the top menu stays behind the taskbar
}

void CMenuContainer::PostRefreshMessage( void )
{
	if (!m_bRefreshPosted && !m_bDestroyed)
	{
		m_bRefreshPosted=true;
		PostMessage(MCM_REFRESH);
	}
}

LRESULT CMenuContainer::OnSysCommand( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled )
{
	if ((wParam&0xFFF0)==SC_KEYMENU)
	{
		// stops Alt from activating the window menu
		ShowKeyboardCues();
	}
	else
		bHandled=FALSE;
	return 0;
}

LRESULT CMenuContainer::OnTimer( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled )
{
	if (wParam==TIMER_HOVER)
	{
		LOG_MENU(LOG_MOUSE,L"End Hover, hover=%d, hot=%d, submenu=%d",m_HoverItem,m_HotItem,m_Submenu);
		// the mouse hovers over an item. open it.
		if (m_HoverItem!=m_Submenu && m_HoverItem==m_HotItem)
			ActivateItem(m_HoverItem,ACTIVATE_OPEN,NULL);
		m_HoverItem=-1;
		KillTimer(TIMER_HOVER);
		return 0;
	}
	if (wParam==TIMER_SCROLL)
	{
		int speed=GetSettingInt(m_bSubMenu?L"SubMenuScrollSpeed":L"MainMenuScrollSpeed");
		if (speed<1) speed=1;
		if (speed>20) speed=20;
		int scroll=m_ScrollOffset;
		if (m_bScrollUp && m_bScrollUpHot)
		{
			scroll-=m_ItemHeight[0]*speed/6;
			if (scroll<0) scroll=0;
		}
		else if (m_bScrollDown && m_bScrollDownHot)
		{
			scroll+=m_ItemHeight[0]*speed/6;
			int total=m_Items[m_ScrollCount-1].itemRect.bottom-m_rContent.top-m_ScrollHeight;
			if (scroll>total) scroll=total;
		}
		if (m_ScrollOffset!=scroll)
		{
			m_ScrollOffset=scroll;
			UpdateScroll();
			if (!m_bScrollUp && !m_bScrollDown)
				KillTimer(TIMER_SCROLL);
			Invalidate();
		}
	}
	if (wParam==TIMER_TOOLTIP_SHOW)
	{
		KillTimer(TIMER_TOOLTIP_SHOW);

		if (!s_pHotMenu || s_pHotMenu->m_bDestroyed)
			return 0;
		if (s_pHotMenu!=s_pTipMenu || s_HotItem!=s_TipItem)
			return 0;

		if (std::find(s_Menus.begin(),s_Menus.end(),s_pHotMenu)==s_Menus.end())
			return 0;

		if (s_HotItem>=(int)s_pHotMenu->m_Items.size())
			return 0;

		TOOLINFO tool={sizeof(tool),TTF_ABSOLUTE|TTF_TRACK|TTF_TRANSPARENT|(s_bRTL?TTF_RTLREADING:0)};
		tool.uId=1;

		wchar_t text[1024];
		if (!s_pHotMenu->GetDescription(s_HotItem,text,_countof(text)))
			return 0;

		RECT rc;
		s_pHotMenu->GetItemRect(s_HotItem,rc);
		s_pHotMenu->ClientToScreen(&rc);
		if (rc.left>rc.right)
		{
			int q=rc.left; rc.left=rc.right; rc.right=q;
		}
		DWORD pos=GetMessagePos();
		POINT pt={(short)LOWORD(pos),(short)HIWORD(pos)};
		if (PtInRect(&rc,pt))
		{
			pt.x+=8;
			pt.y+=16;
		}
		else
		{
			pt.x=(rc.left+rc.right)/2;
			pt.y=rc.bottom;
		}

		tool.lpszText=text;
		s_Tooltip.SendMessage(TTM_UPDATETIPTEXT,0,(LPARAM)&tool);
		s_Tooltip.SendMessage(TTM_TRACKPOSITION,0,MAKELONG(pt.x,pt.y));
		s_Tooltip.SendMessage(TTM_TRACKACTIVATE,TRUE,(LPARAM)&tool);

		// make sure the tooltip is inside the monitor
		MONITORINFO info={sizeof(info)};
		GetMonitorInfo(MonitorFromPoint(pt,MONITOR_DEFAULTTONEAREST),&info);
		s_Tooltip.GetWindowRect(&rc);
		int dx=0, dy=0;
		if (rc.left<info.rcMonitor.left) dx=info.rcMonitor.left-rc.left;
		if (rc.right+dx>info.rcMonitor.right) dx-=rc.right-info.rcMonitor.right;
		if (rc.top<info.rcMonitor.top) dy=info.rcMonitor.top-rc.top;
		if (rc.bottom+dy>info.rcMonitor.bottom) dy-=rc.bottom-info.rcMonitor.bottom;
		if (dx || dy)
			s_Tooltip.SendMessage(TTM_TRACKPOSITION,0,MAKELONG(pt.x+dx,pt.y+dy));

		SetTimer(TIMER_TOOLTIP_HIDE,s_pHotMenu->m_Items[s_HotItem].bFolder?s_TipHideTimeFolder:s_TipHideTime);
		return 0;
	}
	if (wParam==TIMER_TOOLTIP_HIDE)
	{
		TOOLINFO tool={sizeof(tool),TTF_ABSOLUTE|TTF_TRACK|TTF_TRANSPARENT};
		tool.uId=1;
		s_Tooltip.SendMessage(TTM_TRACKACTIVATE,FALSE,(LPARAM)&tool);
		KillTimer(TIMER_TOOLTIP_HIDE);
		return 0;
	}
	if (wParam==TIMER_BALLOON_HIDE)
	{
		TOOLINFO tool={sizeof(tool)};
		tool.uId=1;
		if (s_TooltipBaloon.m_hWnd)
			s_TooltipBaloon.SendMessage(TTM_TRACKACTIVATE,FALSE,(LPARAM)&tool);
		KillTimer(TIMER_BALLOON_HIDE);
	}
	return 0;
}

// Handle right-click and the menu keyboard button
LRESULT CMenuContainer::OnContextMenu( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled )
{
	if (s_bNoContextMenu) return 0;
	POINT pt={(short)LOWORD(lParam),(short)HIWORD(lParam)};
	int index;
	if (pt.x!=-1 || pt.y!=-1)
	{
		POINT pt2=pt;
		ScreenToClient(&pt2);
		index=HitTest(pt2);
	}
	else
	{
		index=m_HotItem;
	}
	if (index<0) return 0;
	ActivateItem(index,ACTIVATE_MENU,&pt);
	return 0;
}

LRESULT CMenuContainer::OnKeyDown( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled )
{
	ShowKeyboardCues();

	if (wParam!=VK_UP && wParam!=VK_DOWN && wParam!=VK_LEFT && wParam!=VK_RIGHT && wParam!=VK_ESCAPE && wParam!=VK_RETURN)
		return TRUE;

	int index=m_HotItem;
	if (index<0) index=-1;
	int n=(int)m_Items.size();

	if (wParam==VK_UP)
	{
		// previous item
		if (index<0)
		{
			int column=(m_bTwoColumns?1:0);
			index=0;
			for (int i=0;i<(int)m_Items.size();i++)
			{
				if (m_Items[i].column==column && m_Items[i].itemRect.bottom>m_Items[i].itemRect.top)
				{
					index=i;
					break;
				}
			}
		}
		// skip separators and the (Empty) item in the top menu if hidden
		while (m_Items[index].id==MENU_SEPARATOR || m_Items[index].itemRect.bottom==m_Items[index].itemRect.top)
			index=(index+n-1)%n;
		index=(index+n-1)%n;
		while (m_Items[index].id==MENU_SEPARATOR || m_Items[index].itemRect.bottom==m_Items[index].itemRect.top)
			index=(index+n-1)%n;
		ActivateItem(index,ACTIVATE_SELECT,NULL);
	}
	if (wParam==VK_DOWN)
	{
		// next item
		while (index>=0 && (m_Items[index].id==MENU_SEPARATOR || m_Items[index].itemRect.bottom==m_Items[index].itemRect.top))
			index=(index+1)%n;
		index=(index+1)%n;
		while (m_Items[index].id==MENU_SEPARATOR || m_Items[index].itemRect.bottom==m_Items[index].itemRect.top)
			index=(index+1)%n;
		ActivateItem(index,ACTIVATE_SELECT,NULL);
	}
	bool bBack=((wParam==VK_LEFT && !s_bRTL) || (wParam==VK_RIGHT && s_bRTL));
	if (wParam==VK_ESCAPE || (bBack && (s_Menus.size()>1 || (s_Menus.size()==1 && m_bSubMenu))))
	{
		// close top menu
		if (!s_Menus[s_Menus.size()-1]->m_bDestroyed)
			s_Menus[s_Menus.size()-1]->PostMessage(WM_CLOSE);
		if (s_Menus.size()>=2 && !s_Menus[s_Menus.size()-2]->m_bDestroyed)
			s_Menus[s_Menus.size()-2]->SetActiveWindow();
		if (s_Menus.size()==1)
		{
			if (m_bSubMenu)
			{
				::SetFocus(g_ProgramsButton);
			}
			else
			{
				// HACK: stops the call to SetActiveWindow(NULL). The correct behavior is to not close the taskbar when Esc is pressed
				s_TaskbarState&=~ABS_AUTOHIDE;
			}
		}
	}
	else if (bBack && s_Menus.size()==1 && m_bTwoColumns && index>=0)
	{
		int column=m_Items[index].column?0:1;
		int y0=(m_Items[index].itemRect.top+m_Items[index].itemRect.bottom)/2;
		int dist=INT_MAX;
		index=-1;
		for (size_t i=0;i<m_Items.size();i++)
		{
			if (m_Items[i].id!=MENU_SEPARATOR && m_Items[i].column==column && m_Items[i].itemRect.bottom>m_Items[i].itemRect.top)
			{
				int y=(m_Items[i].itemRect.top+m_Items[i].itemRect.bottom)/2;
				int d=abs(y-y0);
				if (dist>d)
				{
					index=(int)i;
					dist=d;
				}
			}
		}
		if (index>=0)
			ActivateItem(index,ACTIVATE_SELECT,NULL);
	}

	bool bForward=((wParam==VK_RIGHT && !s_bRTL) || (wParam==VK_LEFT && s_bRTL));
	if (wParam==VK_RETURN || bForward)
	{
		// open submenu
		if (index>=0)
		{
			if (m_Items[index].bFolder)
				ActivateItem(index,ACTIVATE_OPEN_KBD,NULL);
			else if (wParam==VK_RETURN)
				ActivateItem(index,ACTIVATE_EXECUTE,NULL);
			else if (bForward && m_bTwoColumns && m_Items[index].column==0)
			{
				int y0=(m_Items[index].itemRect.top+m_Items[index].itemRect.bottom)/2;
				int dist=INT_MAX;
				index=-1;
				for (size_t i=0;i<m_Items.size();i++)
				{
					if (m_Items[i].id!=MENU_SEPARATOR && m_Items[i].column==1 && m_Items[i].itemRect.bottom>m_Items[i].itemRect.top)
					{
						int y=(m_Items[i].itemRect.top+m_Items[i].itemRect.bottom)/2;
						int d=abs(y-y0);
						if (dist>d)
						{
							index=(int)i;
							dist=d;
						}
					}
				}
				if (index>=0)
					ActivateItem(index,ACTIVATE_SELECT,NULL);
			}
		}
		else if (bForward)
		{
			for (int i=(int)m_Items.size()-1;i>=0;i--)
			{
				if (m_Items[i].id!=MENU_SEPARATOR && m_Items[i].itemRect.bottom>m_Items[i].itemRect.top)
				{
					ActivateItem(i,ACTIVATE_SELECT,NULL);
					break;
				}
			}
		}
	}
	return 0;
}

LRESULT CMenuContainer::OnChar( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled )
{
	if (wParam>=0xD800 && wParam<=0xDBFF)
		return TRUE; // don't support supplementary characters

	// find the current menu item
	int index=m_HotItem;
	if (index<0) index=-1;

	// find the next item with that accelerator
	wchar_t buf[2]={(wchar_t)wParam,0};
	CharUpper(buf);

	int n=(int)m_Items.size();

	int i=1;
	for (;i<=n;i++)
		if (m_Items[(index+2*n+i)%n].accelerator==buf[0])
			break;

	if (i>n) return TRUE; // no item was found

	int first=(index+2*n+i)%n;

	for (i++;i<=n;i++)
		if (m_Items[(index+2*n+i)%n].accelerator==buf[0])
			break;

	if (i>n)
	{
		// exactly 1 item has that accelerator
		if (m_Items[first].bFolder)
			ActivateItem(first,ACTIVATE_OPEN_KBD,NULL);
		else
			ActivateItem(first,ACTIVATE_EXECUTE,NULL);
	}
	else
	{
		// multiple items have the same accelerator. select the next one
		ActivateItem(first,ACTIVATE_SELECT,NULL);
	}

	return 0;
}

LRESULT CMenuContainer::OnDestroy( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled )
{
	LOG_MENU(LOG_OPEN,L"Close Menu, ptr=%p",this);
	if (m_pAccessible)
	{
		NotifyWinEvent(EVENT_SYSTEM_MENUPOPUPEND,m_hWnd,OBJID_CLIENT,CHILDID_SELF);
		m_pAccessible->Reset();
		m_pAccessible->Release();
	}
	RevokeDragDrop(m_hWnd);
	// remember the scroll position
	unsigned int key=CalcFNVHash(m_RegName);
	if (m_ScrollHeight>0)
		s_MenuScrolls[key]=m_ScrollOffset;
	else
		s_MenuScrolls.erase(key);

	if (s_pHotMenu==this)
	{
		s_pHotMenu=NULL;
		s_HotItem=-1;
	}
	if (s_pTipMenu==this)
	{
		s_pTipMenu=NULL;
		TOOLINFO tool={sizeof(tool),TTF_ABSOLUTE|TTF_TRACK|TTF_TRANSPARENT};
		tool.uId=1;
		s_Tooltip.SendMessage(TTM_TRACKACTIVATE,FALSE,(LPARAM)&tool);
	}
	m_bDestroyed=true;
	if (this==s_Menus[0])
	{
		// cleanup when the last menu is closed
		if (s_Theme)
			CloseThemeData(s_Theme);
		s_Theme=NULL;
		if (s_PagerTheme)
			CloseThemeData(s_PagerTheme);
		s_PagerTheme=NULL;
		if (s_Tooltip.m_hWnd)
			s_Tooltip.DestroyWindow();
		s_Tooltip.m_hWnd=NULL;
		s_TooltipBaloon.m_hWnd=NULL; // the balloon tooltip is owned, no need to be destroyed
		s_pHotMenu=NULL;
		s_HotItem=-1;
		if (!m_bSubMenu)
			EnableStartTooltip(true);
		BufferedPaintUnInit();
		FreeMenuSkin(s_Skin);
		if (!m_bSubMenu && (s_TaskbarState&ABS_AUTOHIDE))
			::SetActiveWindow(NULL); // close the taskbar if it is auto-hide
		if (s_XMouse)
			SystemParametersInfo(SPI_SETACTIVEWINDOWTRACKING,NULL,(PVOID)TRUE,SPIF_SENDCHANGE);
		CloseLog();
		s_bAllPrograms=false;
		if ((m_Options&CONTAINER_ALLPROGRAMS) && g_TopMenu && ::IsWindowVisible(g_TopMenu))
		{
			::ShowWindow(g_UserPic,SW_SHOW);
			::SetFocus(g_ProgramsButton);
			CPoint pt(GetMessagePos());
			::ScreenToClient(g_TopMenu,&pt);
			::PostMessage(g_TopMenu,WM_MOUSEMOVE,0,MAKELONG(pt.x,pt.y));
		}
	}
	return 0;
}

LRESULT CMenuContainer::OnRefresh( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled )
{
	// updates the menu after drag/drop, delete, or rename operation
	m_bRefreshPosted=false;
	for (std::vector<CMenuContainer*>::reverse_iterator it=s_Menus.rbegin();*it!=this;++it)
		if (!(*it)->m_bDestroyed)
			(*it)->PostMessage(WM_CLOSE);
	unsigned int key=CalcFNVHash(m_RegName);
	if (m_ScrollHeight>0)
		s_MenuScrolls[key]=m_ScrollOffset;
	else
		s_MenuScrolls.erase(key);
	InitItems();
	InitWindow();
	Invalidate();
	SetFocus();
	SetHotItem(-1);
	return 0;
}

LRESULT CMenuContainer::OnActivate( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled )
{
	if (LOWORD(wParam)!=WA_INACTIVE)
	{
		if (s_Tooltip.m_hWnd)
			s_Tooltip.SetWindowPos(HWND_TOP,0,0,0,0,SWP_NOSIZE|SWP_NOMOVE|SWP_NOACTIVATE);
		return 0;
	}
#ifdef ALLOW_DEACTIVATE
	if (s_pDragSource) return 0;
	// check if another menu window is being activated
	// if not, close all menus
	for (std::vector<CMenuContainer*>::const_iterator it=s_Menus.begin();it!=s_Menus.end();++it)
		if ((*it)->m_hWnd==(HWND)lParam)
			return 0;

	if (g_OwnerWindow && (HWND)lParam==g_OwnerWindow)
		return 0;

	if (g_TopMenu && (HWND)lParam==g_TopMenu)
		return 0;

	for (std::vector<CMenuContainer*>::reverse_iterator it=s_Menus.rbegin();it!=s_Menus.rend();++it)
		if (!(*it)->m_bDestroyed)
			(*it)->PostMessage(WM_CLOSE);
	if (g_TopMenu && s_bAllPrograms) ::PostMessage(g_TopMenu,WM_CLOSE,0,0);
#endif

	return 0;
}

LRESULT CMenuContainer::OnMouseActivate( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled )
{
	if (m_Submenu>=0)
		return MA_NOACTIVATE;
	bHandled=FALSE;
	return 0;
}

LRESULT CMenuContainer::OnMouseMove( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled )
{
	if (!m_bTrackMouse)
	{
		TRACKMOUSEEVENT track={sizeof(track),TME_LEAVE,m_hWnd,0};
		TrackMouseEvent(&track);
		m_bTrackMouse=true;
	}
	if (m_HotPos==GetMessagePos())
		return 0; // HACK - ignore the mouse if it hasn't moved since last time. otherwise the mouse can override the keyboard navigation
	m_HotPos=GetMessagePos();
	POINT pt={(short)LOWORD(lParam),(short)HIWORD(lParam)};
	int index=HitTest(pt);

	if (GetCapture()==m_hWnd)
	{
		if (m_ClickIndex!=index)
		{
			if (!DragOut(m_ClickIndex))
				SetHotItem(-2);
		}
		else
			SetHotItem(index);
	}
	else
	{
		if (index>=0 && m_Items[index].id==MENU_SEPARATOR)
			index=m_HotItem;
		SetHotItem(index);

		UpdateScroll(&pt);

		if (m_Submenu<0)
			SetFocus();
		if (index>=0)
		{
			if ((m_Submenu>=0 && index!=m_Submenu) || (m_Submenu<0 && m_Items[index].bFolder))
			{
				// initialize the hover timer
				if (m_HoverItem!=index)
				{
					m_HoverItem=index;
					SetTimer(TIMER_HOVER,s_HoverTime);
					LOG_MENU(LOG_MOUSE,L"Start Hover, index=%d",index);
				}
			}
			else
				m_HoverItem=-1;
		}
		else
			m_HoverItem=-1;
	}

	return 0;
}

LRESULT CMenuContainer::OnMouseLeave( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled )
{
	UpdateScroll(NULL);
	SetHotItem(-1);
	m_bTrackMouse=false;
	if (m_HoverItem!=-1)
	{
		KillTimer(TIMER_HOVER);
		m_HoverItem=-1;
	}
	return 0;
}

LRESULT CMenuContainer::OnMouseWheel( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled )
{
	POINT pt={(short)LOWORD(lParam),(short)HIWORD(lParam)};
	HWND hwnd=WindowFromPoint(pt);
	if (hwnd!=m_hWnd)
	{
		for (std::vector<CMenuContainer*>::iterator it=s_Menus.begin();it!=s_Menus.end();++it)
		{
			if ((*it)->m_hWnd==hwnd)
			{
				(*it)->SendMessage(uMsg,wParam,lParam);
				return 0;
			}
		}
	}
	if (m_ScrollCount<1) return 0; // nothing to scroll
	UINT lines;
	if (!SystemParametersInfo(SPI_GETWHEELSCROLLLINES,0,&lines,FALSE))
		lines=3;
	if (lines<1) lines=1;

	m_MouseWheel+=lines*(short)HIWORD(wParam);
	int n=m_MouseWheel/WHEEL_DELTA;
	m_MouseWheel-=n*WHEEL_DELTA;
	int scroll=m_ScrollOffset;
	scroll-=n*m_ItemHeight[0];
	if (scroll<0) scroll=0;
	int total=m_Items[m_ScrollCount-1].itemRect.bottom-m_rContent.top-m_ScrollHeight;
	if (scroll>total) scroll=total;
	if (m_ScrollOffset!=scroll)
	{
		m_ScrollOffset=scroll;
		UpdateScroll();
		Invalidate();
		m_HotPos=-1;
		ScreenToClient(&pt);
		OnMouseMove(WM_MOUSEMOVE,LOWORD(wParam),MAKELONG(pt.x,pt.y),bHandled);
	}
	return 0;
}

bool CMenuContainer::GetDescription( int index, wchar_t *text, int size )
{
	if (index<0 || index>=(int)m_Items.size())
		return false;
	const MenuItem &item=m_Items[index];
	if (item.pStdItem && item.pStdItem->tip)
	{
		// get the tip for the standard item
		Strcpy(text,size,item.pStdItem->tip);
		return true;
	}

	// get the tip from the shell
	CComPtr<IShellFolder> pFolder;
	PCUITEMID_CHILD pidl;
	if (FAILED(SHBindToParent(item.pItem1,IID_IShellFolder,(void**)&pFolder,&pidl)))
	return false;

	CComPtr<IQueryInfo> pQueryInfo;
	if (FAILED(pFolder->GetUIObjectOf(NULL,1,&pidl,IID_IQueryInfo,NULL,(void**)&pQueryInfo)))
	return false;

	wchar_t *tip;
	if (FAILED(pQueryInfo->GetInfoTip(QITIPF_DEFAULT,&tip)) || !tip)
	return false;

	Strcpy(text,size,tip);
	CoTaskMemFree(tip);
	return true;
}

LRESULT CMenuContainer::OnLButtonDown( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled )
{
	if (!GetCapture())
	{
		if (m_Submenu<0)
			SetFocus();
		POINT pt={(short)LOWORD(lParam),(short)HIWORD(lParam)};
		m_ClickIndex=-1;
		if (m_rUser1.left<m_rUser1.right && PtInRect(&m_rUser1,pt))
		{
			RunUserCommand(true);
		}
		if (m_rUser2.left<m_rUser2.right && PtInRect(&m_rUser2,pt))
		{
			RunUserCommand(false);
		}
		int index=HitTest(pt);
		if (index<0)
		{
			if (m_Submenu>=0)
			{
				SetActiveWindow(); // must be done before the children are destroyed
				// close all child menus
				for (size_t i=s_Menus.size()-1;s_Menus[i]!=this;i--)
					if (!s_Menus[i]->m_bDestroyed)
						s_Menus[i]->DestroyWindow();
				SetHotItem(-1); // must be done after the children are destroyed
			}
			return 0;
		}
		const MenuItem &item=m_Items[index];
		if (item.id==MENU_SEPARATOR) return 0;
		m_ClickIndex=index;
		SetCapture();
	}
	return 0;
}

LRESULT CMenuContainer::OnLButtonDblClick( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled )
{
	m_ClickIndex=-1;
	// execute item under the mouse
	POINT pt={(short)LOWORD(lParam),(short)HIWORD(lParam)};
	int index=HitTest(pt);
	if (index<0) return 0;
	const MenuItem &item=m_Items[index];
	if (item.id==MENU_SEPARATOR) return 0;
	ClientToScreen(&pt);
	ActivateItem(index,ACTIVATE_EXECUTE,&pt);
	return 0;
}

LRESULT CMenuContainer::OnLButtonUp( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled )
{
	if (GetCapture()!=m_hWnd)
		return 0;
	ReleaseCapture();
	POINT pt={(short)LOWORD(lParam),(short)HIWORD(lParam)};
	int index=HitTest(pt);
	if (index!=m_ClickIndex)
	{
		SetHotItem(-1);
		return 0;
	}
	if (index<0) return 0;
	const MenuItem &item=m_Items[index];
	ClientToScreen(&pt);
	if (!item.bFolder)
		ActivateItem(index,ACTIVATE_EXECUTE,&pt);
	else if (index!=m_Submenu)
		ActivateItem(index,ACTIVATE_OPEN,NULL);
	return 0;
}

LRESULT CMenuContainer::OnRButtonDown( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled )
{
	if (!GetCapture())
	{
		SetFocus();
		POINT pt={(short)LOWORD(lParam),(short)HIWORD(lParam)};
		m_ClickIndex=-1;
		int index=HitTest(pt);
		if (index<0) return 0;
		const MenuItem &item=m_Items[index];
		if (item.id==MENU_SEPARATOR) return 0;
		m_ClickIndex=index;
		SetCapture();
	}
	return 0;
}

LRESULT CMenuContainer::OnRButtonUp( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled )
{
	if (GetCapture()!=m_hWnd)
		return 0;
	ReleaseCapture();
	if (s_bNoContextMenu) return 0;
	POINT pt={(short)LOWORD(lParam),(short)HIWORD(lParam)};
	int index=HitTest(pt);
	if (index<0) return 0;
	const MenuItem &item=m_Items[index];
	if (item.id==MENU_SEPARATOR) return 0;
	ClientToScreen(&pt);
	ActivateItem(index,ACTIVATE_MENU,&pt);
	return 0;
}

LRESULT CMenuContainer::OnSetCursor( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled )
{
	if (m_rUser1.left<m_rUser1.right)
	{
		DWORD pos=GetMessagePos();
		POINT pt={(short)LOWORD(pos),(short)HIWORD(pos)};
		ScreenToClient(&pt);
		if (PtInRect(&m_rUser1,pt))
		{
			SetCursor(LoadCursor(NULL,IDC_HAND));
			return TRUE;
		}
	}
	if (m_rUser2.left<m_rUser2.right)
	{
		DWORD pos=GetMessagePos();
		POINT pt={(short)LOWORD(pos),(short)HIWORD(pos)};
		ScreenToClient(&pt);
		if (PtInRect(&m_rUser2,pt))
		{
			SetCursor(LoadCursor(NULL,IDC_HAND));
			return TRUE;
		}
	}
	bHandled=FALSE;
	return 0;
}

void CMenuContainer::PlayMenuSound( TMenuSound sound )
{
	const wchar_t *setting=NULL;
	switch (sound)
	{
		case SOUND_MAIN:
			setting=L"SoundMain";
			break;
		case SOUND_POPUP:
			setting=L"SoundPopup";
			break;
		case SOUND_COMMAND:
			setting=L"SoundCommand";
			break;
		case SOUND_DROP:
			setting=L"SoundDrop";
			break;
	}
	CString str=GetSettingString(setting);
	if (_wcsicmp(PathFindExtension(str),L".wav")==0)
	{
		wchar_t path[_MAX_PATH];
		Strcpy(path,_countof(path),str);
		DoEnvironmentSubst(path,_countof(path));
		PlaySound(path,NULL,SND_FILENAME|SND_ASYNC|SND_NODEFAULT|SND_SYSTEM);
	}
	else if (_wcsicmp(str,L"none")==0)
		return;
	else
		PlaySound(str,NULL,SND_APPLICATION|SND_ALIAS|SND_ASYNC|SND_NODEFAULT|SND_SYSTEM);
}

void CMenuContainer::SaveItemOrder( const std::vector<SortMenuItem> &items )
{
	if (m_Options&CONTAINER_DROP)
	{
		// save item names in the registry
		CRegKey regOrder;
		if (regOrder.Open(HKEY_CURRENT_USER,m_RegName)!=ERROR_SUCCESS)
			regOrder.Create(HKEY_CURRENT_USER,m_RegName);
		std::vector<unsigned int> hashes;
		hashes.reserve(items.size());
		for (std::vector<SortMenuItem>::const_iterator it=items.begin();it!=items.end();++it)
			hashes.push_back(it->nameHash);
		if (hashes.empty())
			regOrder.SetBinaryValue(L"Order",NULL,0);
		else
			regOrder.SetBinaryValue(L"Order",&hashes[0],(int)hashes.size()*4);
	}
}

void CMenuContainer::LoadItemOrder( void )
{
	bool bLoaded=false;
	if (m_Options&CONTAINER_DROP)
	{
		// load item names from the registry
		std::vector<unsigned int> hashes;
		CRegKey regOrder;
		if (regOrder.Open(HKEY_CURRENT_USER,m_RegName)==ERROR_SUCCESS)
		{
			ULONG size=0;
			regOrder.QueryBinaryValue(L"Order",NULL,&size);
			if (size>0)
			{
				hashes.resize(size/4);
				regOrder.QueryBinaryValue(L"Order",&hashes[0],&size);
				bLoaded=true;
			}
		}
		if (hashes.size()==1 && hashes[0]=='AUTO')
		{
			m_Options|=CONTAINER_AUTOSORT;
			for (std::vector<MenuItem>::iterator it=m_Items.begin();it!=m_Items.end();++it)
				it->row=0;
		}
		else
		{
			m_Options&=~CONTAINER_AUTOSORT;

			unsigned int hash0=CalcFNVHash(L"");

			// assign each m_Item an index based on its position in items. store in row
			// unknown items get the index of the blank item, or at the end
			for (std::vector<MenuItem>::iterator it=m_Items.begin();it!=m_Items.end();++it)
			{
				unsigned int hash=it->nameHash;
				it->row=(int)hashes.size();
				for (size_t i=0;i<hashes.size();i++)
				{
					if (hashes[i]==hash)
					{
						it->row=(int)i;
						break;
					}
					else if (hashes[i]==hash0)
						it->row=(int)i;
				}
				if (m_Options&CONTAINER_SORTZA)
					it->row=-it->row;
			}
		}
	}
	else
	{
		for (std::vector<MenuItem>::iterator it=m_Items.begin();it!=m_Items.end();++it)
			it->row=0;
	}

	// sort by row, then by bFolder, then by name
	std::sort(m_Items.begin(),m_Items.end());
	if (m_Options&CONTAINER_SORTZA)
		std::reverse(m_Items.begin(),m_Items.end());

	if ((m_Options&CONTAINER_DROP) && (m_Options&CONTAINER_SORTONCE) && !bLoaded)
	{
		std::vector<SortMenuItem> items;
		for (std::vector<MenuItem>::const_iterator it=m_Items.begin();it!=m_Items.end();++it)
			if (it->id==MENU_NO)
			{
				SortMenuItem item={it->name,it->nameHash,it->bFolder};
				items.push_back(item);
			}
		SaveItemOrder(items);
	}
}

void CMenuContainer::AddMRUShortcut( const wchar_t *path )
{
	bool bFound=false;
	for (int i=0;i<MRU_PROGRAMS_COUNT;i++)
	{
		if (_wcsicmp(s_MRUShortcuts[i],path)==0)
		{
			if (i>0)
			{
				CString str=s_MRUShortcuts[i];
				for (;i>0;i--)
					s_MRUShortcuts[i]=s_MRUShortcuts[i-1];
				s_MRUShortcuts[0]=str;
			}
			bFound=true;
			break;
		}
	}

	if (!bFound)
	{
		for (int i=MRU_PROGRAMS_COUNT-1;i>0;i--)
			s_MRUShortcuts[i]=s_MRUShortcuts[i-1];
		s_MRUShortcuts[0]=path;
	}

	SaveMRUShortcuts();
}

void CMenuContainer::DeleteMRUShortcut( const wchar_t *path )
{
	for (int i=0;i<MRU_PROGRAMS_COUNT;i++)
	{
		if (_wcsicmp(s_MRUShortcuts[i],path)==0)
		{
			for (;i<MRU_PROGRAMS_COUNT-1;i++)
				s_MRUShortcuts[i]=s_MRUShortcuts[i+1];
			break;
		}
	}

	SaveMRUShortcuts();
}

void CMenuContainer::SaveMRUShortcuts( void )
{
	CRegKey regMRU;
	if (regMRU.Open(HKEY_CURRENT_USER,L"Software\\IvoSoft\\ClassicStartMenu\\MRU")!=ERROR_SUCCESS)
	regMRU.Create(HKEY_CURRENT_USER,L"Software\\IvoSoft\\ClassicStartMenu\\MRU");

	bool bDelete=false;
	for (int i=0;i<MRU_PROGRAMS_COUNT;i++)
	{
		wchar_t name[10];
		Sprintf(name,_countof(name),L"%d",i);
		if (s_MRUShortcuts[i].IsEmpty())
			bDelete=true; // delete the rest!
		if (bDelete)
		{
			wchar_t path[256];
			ULONG size=_countof(path);
			if (regMRU.QueryStringValue(name,path,&size)!=ERROR_SUCCESS)
				break;
			regMRU.DeleteValue(name);
		}
		else
			regMRU.SetStringValue(name,s_MRUShortcuts[i]);
	}
}

void CMenuContainer::LoadMRUShortcuts( void )
{
	for (int i=0;i<MRU_PROGRAMS_COUNT;i++)
		s_MRUShortcuts[i].Empty();
	CRegKey regMRU;
	if (regMRU.Open(HKEY_CURRENT_USER,L"Software\\IvoSoft\\ClassicStartMenu\\MRU")==ERROR_SUCCESS)
	{
		for (int i=0;i<MRU_PROGRAMS_COUNT;i++)
		{
			wchar_t name[10];
			Sprintf(name,_countof(name),L"%d",i);
			wchar_t path[256];
			ULONG size=_countof(path);
			if (regMRU.QueryStringValue(name,path,&size)!=ERROR_SUCCESS)
				break;
			s_MRUShortcuts[i]=path;
		}
	}
}

LRESULT CMenuContainer::OnGetAccObject( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled )
{
	if ((DWORD)lParam==(DWORD)OBJID_CLIENT && m_pAccessible)
	{
		return LresultFromObject(IID_IAccessible,wParam,m_pAccessible);
	}
	else
	{
		bHandled=FALSE;
		return 0;
	}
}

///////////////////////////////////////////////////////////////////////////////

bool CMenuContainer::CloseStartMenu( void )
{
	if (s_Menus.empty()) return false;

	::SetActiveWindow(g_StartButton);

	if (s_LastFGWindow)
		SetForegroundWindow(s_LastFGWindow);

	return true;
}

void CMenuContainer::HideStartMenu( void )
{
	for (std::vector<CMenuContainer*>::iterator it=s_Menus.begin();it!=s_Menus.end();++it)
		if (!(*it)->m_bDestroyed)
			(*it)->ShowWindow(SW_HIDE);
}

bool CMenuContainer::CloseProgramsMenu( void )
{
	if (s_Menus.empty()) return false;

	for (std::vector<CMenuContainer*>::const_reverse_iterator it=s_Menus.rbegin();it!=s_Menus.rend();++it)
		if (!(*it)->m_bDestroyed)
			(*it)->PostMessage(WM_CLOSE);

	return true;
}

// Toggles the start menu
HWND CMenuContainer::ToggleStartMenu( HWND startButton, bool bKeyboard, bool bAllPrograms )
{
	s_bAllPrograms=false;
	if (bAllPrograms)
	{
		if (CloseProgramsMenu())
		{
			return NULL;
		}

		::ShowWindow(g_UserPic,SW_HIDE);
	}
	else
	{
		if (!bKeyboard) s_LastFGWindow=NULL;
		if (CloseStartMenu())
			return NULL;

		s_LastFGWindow=GetForegroundWindow();
		SetForegroundWindow(g_StartButton);
		EnableStartTooltip(false);
	}

	s_bAllPrograms=bAllPrograms;
	int level=GetSettingInt(L"LogLevel");
	if (level>0)
	{
		wchar_t fname[_MAX_PATH]=L"%LOCALAPPDATA%\\StartMenuLog.txt";
		DoEnvironmentSubst(fname,_countof(fname));
		InitLog((TLogLevel)level,fname);
	}

	{
		CSettingsLockWrite lock;
		UpdateDefaultSettings();
	}

	// initialize all settings
	bool bErr=false;
	if (bAllPrograms)
	{
		bErr=!LoadMenuSkin(GetSettingString(L"Skin2"),s_Skin,GetSettingString(L"SkinVariation2"),GetSettingString(L"SkinOptions2"),LOADMENU_RESOURCES);
	}
	else
		bErr=!LoadMenuSkin(GetSettingString(L"Skin1"),s_Skin,GetSettingString(L"SkinVariation1"),GetSettingString(L"SkinOptions1"),LOADMENU_RESOURCES|LOADMENU_MAIN);
	if (bErr)
		LoadDefaultMenuSkin(s_Skin,LOADMENU_MAIN|LOADMENU_RESOURCES);

	s_ScrollMenus=GetSettingInt(L"ScrollType");
	s_bExpandLinks=GetSettingBool(L"ExpandFolderLinks");
	s_bLogicalSort=GetSettingBool(L"NumericSort");
	s_MaxRecentDocuments=GetSettingInt(L"MaxRecentDocuments");
	s_ShellFormat=RegisterClipboardFormat(CFSTR_SHELLIDLIST);

	bool bRemote=GetSystemMetrics(SM_REMOTESESSION)!=0;
	wchar_t wabPath[_MAX_PATH]=L"%ProgramFiles%\\Windows Mail\\wab.exe";
	DoEnvironmentSubst(wabPath,_countof(wabPath));
	HANDLE hWab=CreateFile(wabPath,0,FILE_SHARE_READ,NULL,OPEN_EXISTING,FILE_ATTRIBUTE_NORMAL,NULL);
	bool bPeople=(hWab!=INVALID_HANDLE_VALUE);
	if (bPeople) CloseHandle(hWab);
	s_bRTL=s_Skin.ForceRTL || IsLanguageRTL();

	APPBARDATA appbar={sizeof(appbar)};
	s_TaskbarState=(DWORD)SHAppBarMessage(ABM_GETSTATE,&appbar);
	DWORD version=LOWORD(GetVersion());
	// the taskbar on Windows 7 (and most likely later versions) is always on top even though it doesn't have the ABS_ALWAYSONTOP flag.
	// to check the version we have to swap the low and high bytes returned by GetVersion(). Would have been so much easier if GetVersion
	// returns the bytes in the correct order rather than being retarded like this.
	if (MAKEWORD(HIBYTE(version),LOBYTE(version))>=0x601)
		s_TaskbarState|=ABS_ALWAYSONTOP;

	appbar.hWnd=g_TaskBar;
	// get the taskbar orientation - top/bottom/left/right. the documentation says only the bounding rectangle (rc) is returned, but seems
	// that also uEdge is returned. hopefully this won't break in the future. it is tricky to calculate the position of the taskbar only
	// based on the bounding rectangle. there may be other appbars docked on each side, the taskbar can be set to auto-hide, etc. tricky.
	SHAppBarMessage(ABM_GETTASKBARPOS,&appbar);

	if (!bAllPrograms && (s_TaskbarState&ABS_AUTOHIDE))
		::SetActiveWindow(g_TaskBar);

	if (s_bActiveDirectory==-1)
	{
		DOMAIN_CONTROLLER_INFO *info;
		DWORD err=DsGetDcName(NULL,NULL,NULL,NULL,DS_RETURN_FLAT_NAME,&info);
		if (err==ERROR_SUCCESS)
		{
			s_bActiveDirectory=1;
			NetApiBufferFree(info);
		}
		else
			s_bActiveDirectory=0;
	}

	SYSTEM_POWER_CAPABILITIES powerCaps;
	GetPwrCapabilities(&powerCaps);

	for (int i=0;i<_countof(g_StdOptions);i++)
	{
		switch (g_StdOptions[i].id)
		{
			case MENU_FAVORITES:
				g_StdOptions[i].options=0;
				{
					int show=GetSettingInt(L"Favorites");
					if (show==1)
						g_StdOptions[i].options=MENU_ENABLED;
					else if (show==2)
						g_StdOptions[i].options=MENU_ENABLED|MENU_EXPANDED;
				}
				break;
			case MENU_DOCUMENTS:
				g_StdOptions[i].options=0;
				{
					int show=GetSettingInt(L"Documents");
					if (show==1)
						g_StdOptions[i].options=MENU_ENABLED;
					else if (show==2)
						g_StdOptions[i].options=MENU_ENABLED|MENU_EXPANDED;
				}
				break;
			case MENU_LOGOFF:
				g_StdOptions[i].options=GetSettingBool(L"LogOff")?MENU_ENABLED|MENU_EXPANDED:0;
				break;
			case MENU_DISCONNECT:
				g_StdOptions[i].options=(bRemote && !SHRestricted(REST_NODISCONNECT))?MENU_ENABLED|MENU_EXPANDED:0;
				break;
			case MENU_SHUTDOWN:
				g_StdOptions[i].options=(!bRemote || GetSettingBool(L"RemoteShutdown"))?MENU_ENABLED|MENU_EXPANDED:0;
				break;
			case MENU_SHUTDOWN_BOX:
				g_StdOptions[i].options=0;
				if (!bRemote || GetSettingBool(L"RemoteShutdown"))
				{
					int show=GetSettingInt(L"Shutdown");
					if (show==1)
						g_StdOptions[i].options=MENU_ENABLED;
					else if (show==2)
						g_StdOptions[i].options=MENU_ENABLED|MENU_EXPANDED;
				}
				break;
			case MENU_UNDOCK:
				{
					HW_PROFILE_INFO info;
					GetCurrentHwProfile(&info);
					g_StdOptions[i].options=(((info.dwDockInfo&(DOCKINFO_DOCKED|DOCKINFO_UNDOCKED))==DOCKINFO_DOCKED) && GetSettingBool(L"Undock"))?MENU_ENABLED|MENU_EXPANDED:0;
				}
				break;
			case MENU_CONTROLPANEL:
				g_StdOptions[i].options=0;
				{
					int show=GetSettingInt(L"ControlPanel");
					if (show==1)
						g_StdOptions[i].options=MENU_ENABLED;
					else if (show==2)
						g_StdOptions[i].options=MENU_ENABLED|MENU_EXPANDED;
				}
				break;
			case MENU_SECURITY:
				g_StdOptions[i].options=bRemote?MENU_ENABLED:0;
				break;
			case MENU_NETWORK:
				g_StdOptions[i].options=0;
				{
					int show=GetSettingInt(L"Network");
					if (show==1)
						g_StdOptions[i].options=MENU_ENABLED;
					else if (show==2)
						g_StdOptions[i].options=MENU_ENABLED|MENU_EXPANDED;
				}
				break;
			case MENU_PRINTERS:
				g_StdOptions[i].options=0;
				{
					int show=GetSettingInt(L"Printers");
					if (show==1)
						g_StdOptions[i].options=MENU_ENABLED;
					else if (show==2)
						g_StdOptions[i].options=MENU_ENABLED|MENU_EXPANDED;
				}
				break;

			case MENU_SEARCH_PRINTER:
				g_StdOptions[i].options=s_bActiveDirectory==1?MENU_ENABLED|MENU_EXPANDED:0;
				break;
			case MENU_SEARCH_COMPUTERS:
				g_StdOptions[i].options=(s_bActiveDirectory==1 && !SHRestricted(REST_HASFINDCOMPUTERS))?MENU_ENABLED|MENU_EXPANDED:0;
				break;
			case MENU_SEARCH_FILES:
				{
					bool bEnabled=false;
					if (!GetSettingString(L"SearchFilesCommand").IsEmpty())
						bEnabled=true;
					else
					{
						CComPtr<IShellDispatch2> pShellDisp;
						if (SUCCEEDED(CoCreateInstance(CLSID_Shell,NULL,CLSCTX_SERVER,IID_IShellDispatch2,(void**)&pShellDisp)))
						{
							CComVariant var;
							if (SUCCEEDED(pShellDisp->IsServiceRunning(CComBSTR(L"WSearch"),&var)) && var.vt==VT_BOOL && var.boolVal)
								bEnabled=true;
						}
					}
					g_StdOptions[i].options=bEnabled?MENU_ENABLED|MENU_EXPANDED:0;
				}
				break;
			case MENU_SEARCH_PEOPLE:
				g_StdOptions[i].options=bPeople?MENU_ENABLED|MENU_EXPANDED:0;
				break;

			case MENU_HELP:
				g_StdOptions[i].options=GetSettingBool(L"Help")?MENU_ENABLED|MENU_EXPANDED:0;
				break;
			case MENU_RUN:
				g_StdOptions[i].options=GetSettingBool(L"Run")?MENU_ENABLED|MENU_EXPANDED:0;
				break;
			case MENU_TASKBAR:
				g_StdOptions[i].options=!SHRestricted(REST_NOSETTASKBAR)?MENU_ENABLED|MENU_EXPANDED:0;
				break;
			case MENU_FEATURES:
				g_StdOptions[i].options=(!SHRestricted(REST_NOSETFOLDERS) && !SHRestricted(REST_NOCONTROLPANEL))?MENU_ENABLED|MENU_EXPANDED:0;
				break;
			case MENU_CLASSIC_SETTINGS:
				g_StdOptions[i].options=GetSettingBool(L"EnableSettings")?MENU_ENABLED|MENU_EXPANDED:0;
				break;
			case MENU_SEARCH:
				g_StdOptions[i].options=GetSettingBool(L"Search")?MENU_ENABLED|MENU_EXPANDED:0;
				break;
			case MENU_USERFILES:
				g_StdOptions[i].options=0;
				{
					int show=GetSettingInt(L"UserFiles");
					if (show==1)
						g_StdOptions[i].options=MENU_ENABLED;
					else if (show==2)
						g_StdOptions[i].options=MENU_ENABLED|MENU_EXPANDED;
				}
				break;
			case MENU_USERDOCUMENTS:
				g_StdOptions[i].options=0;
				{
					int show=GetSettingInt(L"UserDocuments");
					if (show==1)
						g_StdOptions[i].options=MENU_ENABLED;
					else if (show==2)
						g_StdOptions[i].options=MENU_ENABLED|MENU_EXPANDED;
				}
				break;
			case MENU_USERPICTURES:
				g_StdOptions[i].options=0;
				{
					int show=GetSettingInt(L"UserPictures");
					if (show==1)
						g_StdOptions[i].options=MENU_ENABLED;
					else if (show==2)
						g_StdOptions[i].options=MENU_ENABLED|MENU_EXPANDED;
				}
				break;
			case MENU_SLEEP:
				g_StdOptions[i].options=powerCaps.SystemS3?MENU_ENABLED|MENU_EXPANDED:0;
				break;
			case MENU_HIBERNATE:
				g_StdOptions[i].options=powerCaps.HiberFilePresent?MENU_ENABLED|MENU_EXPANDED:0;
				break;
		}
		LOG_MENU(LOG_OPEN,L"ItemOptions[%d]=%d",i,g_StdOptions[i].options);
	}	

	s_bNoDragDrop=!GetSettingBool(L"EnableDragDrop");
	s_bNoContextMenu=!GetSettingBool(L"EnableContextMenu");
	s_bKeyboardCues=bKeyboard;
	s_bRecentItems=GetSettingBool(L"RecentPrograms");

	// make sure the menu stays on the same monitor as the start button
	RECT startRect;
	::GetWindowRect(startButton,&startRect);
	MONITORINFO info={sizeof(info)};
	GetMonitorInfo(MonitorFromWindow(startButton,MONITOR_DEFAULTTOPRIMARY),&info);
	s_MainRect=info.rcMonitor;

	RECT taskbarRect;
	::GetWindowRect(g_TaskBar,&taskbarRect);
	s_HoverTime=GetSettingInt(L"MenuDelay");

	s_TipShowTime=400;
	s_TipHideTime=4000;
	CString str=GetSettingString(L"InfotipDelay");
	if (!str.IsEmpty())
	{
		wchar_t token[256];
		str=GetToken(str,token,_countof(token),L", \t");
		int time=_wtol(token);
		if (time>=0) s_TipShowTime=time;
		str=GetToken(str,token,_countof(token),L", \t");
		time=_wtol(token);
		if (time>=0) s_TipHideTime=time;
	}

	s_TipHideTimeFolder=s_TipShowTimeFolder=0;
	str=GetSettingString(L"FolderInfotipDelay");
	if (!str.IsEmpty())
	{
		wchar_t token[256];
		str=GetToken(str,token,_countof(token),L", \t");
		int time=_wtol(token);
		if (time>=0) s_TipShowTimeFolder=time;
		str=GetToken(str,token,_countof(token),L", \t");
		time=_wtol(token);
		if (time>=0) s_TipHideTimeFolder=time;
	}

	// create the top menu from the Start Menu folders
	PIDLIST_ABSOLUTE path1;
	PIDLIST_ABSOLUTE path2;
	SHGetKnownFolderIDList(FOLDERID_StartMenu,0,NULL,&path1);
	SHGetKnownFolderIDList(FOLDERID_CommonStartMenu,0,NULL,&path2);

	int options=CONTAINER_PROGRAMS|CONTAINER_DRAG|CONTAINER_DROP;
	if (s_Skin.Main_large_icons && !bAllPrograms)
		options|=CONTAINER_LARGE;

	DWORD animFlags=0;
	{
		int anim=GetSettingInt(bAllPrograms?L"SubMenuAnimation":L"MainMenuAnimation");
		if (anim==3) animFlags=((rand()<RAND_MAX/2)?AW_BLEND:AW_SLIDE);
		else if (anim==1) animFlags=AW_BLEND;
		else if (anim==2) animFlags=AW_SLIDE;
	}

	s_bBehindTaskbar=!bAllPrograms;
	s_bShowTopEmpty=false;
	DWORD dwStyle=WS_POPUP;
	s_SubmenuStyle=WS_POPUP;
	bool bTheme=IsAppThemed()!=FALSE;
	if (bTheme)
	{
		BOOL comp;
		if (FAILED(DwmIsCompositionEnabled(&comp)))
			comp=FALSE;

		if (s_Skin.Main_opacity==MenuSkin::OPACITY_SOLID)
			dwStyle|=WS_BORDER;
		else if (!comp)
			s_Skin.Main_opacity=MenuSkin::OPACITY_REGION;

		if (s_Skin.Submenu_opacity==MenuSkin::OPACITY_SOLID)
			s_SubmenuStyle|=WS_BORDER;
		else if (!comp)
			s_Skin.Submenu_opacity=MenuSkin::OPACITY_REGION;
	}
	else
	{
		if (s_Skin.Main_opacity==MenuSkin::OPACITY_SOLID)
			dwStyle|=s_Skin.Main_thin_frame?WS_BORDER:WS_DLGFRAME;
		else
			s_Skin.Main_opacity=MenuSkin::OPACITY_REGION;

		if (s_Skin.Submenu_opacity==MenuSkin::OPACITY_SOLID)
			s_SubmenuStyle|=s_Skin.Submenu_thin_frame?WS_BORDER:WS_DLGFRAME;
		else
			s_Skin.Submenu_opacity=MenuSkin::OPACITY_REGION;
	}

	RECT margin={0,0,0,0};
	AdjustWindowRect(&margin,s_SubmenuStyle,FALSE);
	s_Skin.Submenu_padding.left+=margin.left; if (s_Skin.Submenu_padding.left<0) s_Skin.Submenu_padding.left=0;
	s_Skin.Submenu_padding.right-=margin.right; if (s_Skin.Submenu_padding.right<0) s_Skin.Submenu_padding.right=0;
	s_Skin.Submenu_padding.top+=margin.top; if (s_Skin.Submenu_padding.top<0) s_Skin.Submenu_padding.top=0;
	s_Skin.Submenu_padding.bottom-=margin.bottom; if (s_Skin.Submenu_padding.bottom<0) s_Skin.Submenu_padding.bottom=0;

	RECT rc;
	const StdMenuItem *pRoot=NULL;

	if (bAllPrograms)
	{
		if (!::GetWindowRect(g_ProgramsButton,&rc))
			memset(&rc,0,sizeof(rc));
		if (startRect.right+startRect.left<s_MainRect.left+s_MainRect.right)
		{
			// start button on the left
			options|=CONTAINER_LEFT;
			rc.left=rc.right=rc.right-s_Skin.Submenu_padding.left+s_Skin.AllPrograms_offset;
			s_bExpandRight=true;
		}
		else
		{
			// start button on the right
			s_bExpandRight=false;
			rc.left=rc.right=rc.left+s_Skin.Submenu_padding.right-s_Skin.AllPrograms_offset;
		}
		rc.top=rc.bottom;
		options|=CONTAINER_ALLPROGRAMS;
		options|=CONTAINER_MULTICOLUMN;
	}
	else
	{
		memset(&margin,0,sizeof(margin));
		AdjustWindowRect(&margin,dwStyle,FALSE);
		if (s_Skin.Main_bitmap_slices_X[1]>0)
		{
			s_Skin.Caption_padding.left+=margin.left; if (s_Skin.Caption_padding.left<0) s_Skin.Caption_padding.left=0;
			s_Skin.Caption_padding.top+=margin.top; if (s_Skin.Caption_padding.top<0) s_Skin.Caption_padding.top=0;
			s_Skin.Caption_padding.bottom-=margin.bottom; if (s_Skin.Caption_padding.bottom<0) s_Skin.Caption_padding.bottom=0;
		}
		else
		{
			// no caption
			s_Skin.Main_padding.left+=margin.left; if (s_Skin.Main_padding.left<0) s_Skin.Main_padding.left=0;
			if (s_Skin.Main_padding2.left>=0)
			{
				s_Skin.Main_padding2.left+=margin.left; if (s_Skin.Main_padding2.left<0) s_Skin.Main_padding2.left=0;
			}
		}
		s_Skin.Main_padding.right-=margin.right; if (s_Skin.Main_padding.right<0) s_Skin.Main_padding.right=0;
		s_Skin.Main_padding.top+=margin.top; if (s_Skin.Main_padding.top<0) s_Skin.Main_padding.top=0;
		s_Skin.Main_padding.bottom-=margin.bottom; if (s_Skin.Main_padding.bottom<0) s_Skin.Main_padding.bottom=0;
		if (s_Skin.Main_padding2.left>=0)
		{
			s_Skin.Main_padding2.right-=margin.right; if (s_Skin.Main_padding2.right<0) s_Skin.Main_padding2.right=0;
			s_Skin.Main_padding2.top+=margin.top; if (s_Skin.Main_padding2.top<0) s_Skin.Main_padding2.top=0;
			s_Skin.Main_padding2.bottom-=margin.bottom; if (s_Skin.Main_padding2.bottom<0) s_Skin.Main_padding2.bottom=0;
		}

		if (!bTheme)
			memset(&margin,0,sizeof(margin)); // in Classic mode don't offset the main menu by the border size

		if ((appbar.uEdge==ABE_LEFT || appbar.uEdge==ABE_RIGHT) && GetSettingBool(L"ShowNextToTaskbar"))
		{
			// when the taskbar is on the side and the menu is not on top of it
			// the start button is assumed at the top
			rc.top=rc.bottom=s_MainRect.top+margin.top;
			if (appbar.uEdge==ABE_LEFT)
			{
				rc.left=rc.right=taskbarRect.right+margin.left;
				options|=CONTAINER_LEFT;
			}
			else
			{
				rc.left=rc.right=taskbarRect.left+margin.right;
			}
			options|=CONTAINER_TOP;
			animFlags|=AW_VER_POSITIVE;
			s_bBehindTaskbar=true;
		}
		else
		{
			if (appbar.uEdge==ABE_BOTTOM)
			{
				// taskbar is at the bottom
				rc.top=rc.bottom=taskbarRect.top+margin.bottom;

				// animate up
				animFlags|=AW_VER_NEGATIVE;
			}
			else if (appbar.uEdge==ABE_TOP)
			{
				// taskbar is at the top
				rc.top=rc.bottom=taskbarRect.bottom+margin.top;

				// animate down
				animFlags|=AW_VER_POSITIVE;
				options|=CONTAINER_TOP;
			}
			else
			{
				// taskbar is on the side, start button must be at the top
				rc.top=rc.bottom=startRect.bottom;

				// animate down
				animFlags|=AW_VER_POSITIVE;
				options|=CONTAINER_TOP;
				s_bBehindTaskbar=false;
			}

			if (startRect.right+startRect.left<s_MainRect.left+s_MainRect.right)
			{
				// start button on the left
				rc.left=rc.right=s_MainRect.left+margin.left;
				options|=CONTAINER_LEFT;
				s_bExpandRight=true;
			}
			else
			{
				// start button on the right
				rc.left=rc.right=s_MainRect.right+margin.right;
				s_bExpandRight=false;
			}
		}

		unsigned int rootSettings;
		pRoot=ParseCustomMenu(rootSettings);
		if (GetSettingBool(L"MainSortZA")) options|=CONTAINER_SORTZA;
		if (GetSettingBool(L"MainSortOnce")) options|=CONTAINER_SORTONCE;
		if (s_bRecentItems && !(rootSettings&StdMenuItem::MENU_NORECENT))
			options|=CONTAINER_RECENT;
	}

	CMenuContainer *pStartMenu=new CMenuContainer(NULL,bAllPrograms?0:-1,options,pRoot,path1,path2,bAllPrograms?L"Software\\IvoSoft\\ClassicStartMenu\\Order2":L"Software\\IvoSoft\\ClassicStartMenu\\Order");
	pStartMenu->InitItems();
	pStartMenu->m_MaxWidth=s_MainRect.right-s_MainRect.left;
	ILFree(path1);
	ILFree(path2);

	bool bTopMost=(s_TaskbarState&ABS_ALWAYSONTOP)!=0 || bAllPrograms;

	SystemParametersInfo(SPI_GETACTIVEWINDOWTRACKING,NULL,&s_XMouse,0);
	if (s_XMouse)
		SystemParametersInfo(SPI_SETACTIVEWINDOWTRACKING,NULL,(PVOID)FALSE,SPIF_SENDCHANGE);

	HWND owner=NULL;
	if (bAllPrograms && (appbar.uEdge==ABE_LEFT || appbar.uEdge==ABE_RIGHT))
		owner=g_TopMenu;
	if (!pStartMenu->Create(owner,&rc,NULL,bAllPrograms?s_SubmenuStyle:dwStyle,WS_EX_TOOLWINDOW|(bTopMost?WS_EX_TOPMOST:0)|(s_bRTL?WS_EX_LAYOUTRTL:0)))
	{
		FreeMenuSkin(s_Skin);
		return NULL;
	}
	pStartMenu->SetHotItem((bKeyboard && bAllPrograms)?0:-1);

	if (bAllPrograms)
	{
		::InvalidateRect(g_ProgramsButton,NULL,TRUE);
		::UpdateWindow(g_ProgramsButton);
	}

	BOOL animate;
	if ((animFlags&(AW_BLEND|AW_SLIDE))==0)
		animate=FALSE;
	else
		SystemParametersInfo(SPI_GETMENUANIMATION,NULL,&animate,0);

	if (s_bBehindTaskbar)
		::SetWindowPos(startButton,bTopMost?HWND_TOPMOST:HWND_TOP,0,0,0,0,SWP_NOMOVE|SWP_NOSIZE); // bring the start button on top
	if (animate)
	{
		int speed=GetSettingInt(bAllPrograms?L"SubMenuAnimationSpeed":L"MainMenuAnimationSpeed");
		if (speed<=0) speed=MENU_ANIM_SPEED;
		else if (speed>=10000) speed=10000;
		if (!AnimateWindow(pStartMenu->m_hWnd,speed,animFlags))
			pStartMenu->ShowWindow(SW_SHOW); // show the menu anyway if AnimateWindow fails
	}
	else
		pStartMenu->ShowWindow(SW_SHOW);
	pStartMenu->SetFocus();
	if (!bAllPrograms)
		pStartMenu->SetHotItem(-1);
	SetForegroundWindow(pStartMenu->m_hWnd);
	SwitchToThisWindow(pStartMenu->m_hWnd,FALSE); // just in case
	// position the start button on top
	if (s_bBehindTaskbar)
		::SetWindowPos(startButton,bTopMost?HWND_TOPMOST:HWND_TOP,0,0,0,0,SWP_NOMOVE|SWP_NOSIZE|SWP_NOACTIVATE);

	if (bErr && GetSettingBool(L"ReportSkinErrors") && !*MenuSkin::s_SkinError)
	{
		Strcpy(MenuSkin::s_SkinError,_countof(MenuSkin::s_SkinError),LoadStringEx(IDS_SKIN_ERR_UNKNOWN));
	}
	if (*MenuSkin::s_SkinError && GetSettingBool(L"ReportSkinErrors"))
	{
		Strcat(MenuSkin::s_SkinError,_countof(MenuSkin::s_SkinError),LoadStringEx(IDS_SKIN_ERR_DISABLE));
		s_TooltipBaloon=CreateWindowEx(WS_EX_TOPMOST|WS_EX_TOOLWINDOW|(s_bRTL?WS_EX_LAYOUTRTL:0),TOOLTIPS_CLASS,NULL,WS_POPUP|TTS_BALLOON|TTS_CLOSE|TTS_NOPREFIX,0,0,0,0,pStartMenu->m_hWnd,NULL,g_Instance,NULL);
		s_TooltipBaloon.SendMessage(TTM_SETMAXTIPWIDTH,0,500);
		TOOLINFO tool={sizeof(tool),TTF_TRANSPARENT|TTF_TRACK|(s_bRTL?TTF_RTLREADING:0)};
		tool.uId=1;
		tool.lpszText=MenuSkin::s_SkinError;
		s_TooltipBaloon.SendMessage(TTM_ADDTOOL,0,(LPARAM)&tool);
		if (bErr)
		{
			s_TooltipBaloon.SendMessage(TTM_SETTITLE,TTI_ERROR,(LPARAM)(const wchar_t*)LoadStringEx(IDS_SKIN_ERR));
		}
		else
		{
			s_TooltipBaloon.SendMessage(TTM_SETTITLE,TTI_WARNING,(LPARAM)(const wchar_t*)LoadStringEx(IDS_SKIN_WARN));
		}
		::GetWindowRect(g_StartButton,&rc);
		s_TooltipBaloon.SendMessage(TTM_TRACKPOSITION,0,MAKELONG((rc.left+rc.right)/2,(rc.top+rc.bottom)/2));
		s_TooltipBaloon.SendMessage(TTM_TRACKACTIVATE,TRUE,(LPARAM)&tool);
		pStartMenu->SetTimer(TIMER_BALLOON_HIDE,10000);
	}

	return pStartMenu->m_hWnd;
}

bool CMenuContainer::ProcessMouseMessage( HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam )
{
	if (uMsg==WM_MOUSEMOVE)
	{
		if (!s_bAllPrograms)
			return false;
		if (hwnd==g_ProgramsButton)
			return true;
		for (std::vector<CMenuContainer*>::const_iterator it=s_Menus.begin();it!=s_Menus.end();++it)
			if ((*it)->m_hWnd==hwnd && (*it)->m_ContextItem<0)
			{
				(*it)->SendMessage(WM_MOUSEMOVE,wParam,lParam);
				return true;
			}
	}
	if (uMsg==WM_MOUSELEAVE)
	{
		if (hwnd==g_ProgramsButton)
			return s_bAllPrograms;
	}
	if (uMsg==WM_MOUSEHOVER)
	{
		if (hwnd==g_ProgramsButton && GetSettingBool(L"CascadeAll"))
			return true;
		if (!s_bAllPrograms)
			return false;
		for (std::vector<CMenuContainer*>::const_iterator it=s_Menus.begin();it!=s_Menus.end();++it)
			if ((*it)->m_hWnd==hwnd)
				return false;
		// if the mouse hovers over some window, close the menus
		while (!s_Menus.empty())
			s_Menus[0]->DestroyWindow();
		::ShowWindow(g_UserPic,SW_SHOW);
		CPoint pt(GetMessagePos());
		::ScreenToClient(g_TopMenu,&pt);
		::PostMessage(g_TopMenu,WM_MOUSEMOVE,0,MAKELONG(pt.x,pt.y));
		return false;
	}
	return false;
}
