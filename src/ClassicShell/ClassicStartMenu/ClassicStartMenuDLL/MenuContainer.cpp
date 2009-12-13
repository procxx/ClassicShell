// Classic Shell (c) 2009, Ivo Beltchev
// The sources for Classic Shell are distributed under the MIT open source license

#include "stdafx.h"
#include "IconManager.h"
#include "MenuContainer.h"
#include "ClassicStartMenuDLL.h"
#include "Settings.h"
#include "ParseSettings.h"
#include "FNVHash.h"
#include "resource.h"
#include <uxtheme.h>
#include <vsstyle.h>
#include <vssym32.h>
#include <WtsApi32.h>
#include <Lm.h>
#include <Dsgetdc.h>
#define SECURITY_WIN32
#include <Security.h>
#include <algorithm>

#define ALLOW_DEACTIVATE // undefine this to prevent the menu from closing when it is deactivated (useful for debugging)

#if defined(BUILD_SETUP) && !defined(ALLOW_DEACTIVATE)
#define ALLOW_DEACTIVATE // make sure it is defined in Setup
#endif

const int MAX_MENU_ITEMS=500;
const int MENU_ANIM_SPEED=200;
const int MENU_ANIM_SPEED_SUBMENU=100;

struct StdMenuItem
{
	TMenuID id;
	const char *key; // localization key
	const wchar_t *name2; // default name
	int icon; // index in shell32.dll
	TMenuID submenu; // MENU_NO if no submenu
	const KNOWNFOLDERID *folder1; // NULL if not used
	const KNOWNFOLDERID *folder2; // NULL if not used
	const char *tipKey; // localization key for the tooltip
	const wchar_t *tip2; // default tooltip
};

struct StdMenuOption
{
	TMenuID id;
	int options;
};

// This table defines the standard menu items
static const StdMenuItem g_StdMenu[]=
{
	// Start menu
	{MENU_PROGRAMS,"Menu.Programs",L"&Programs",326,MENU_NO,&FOLDERID_Programs,&FOLDERID_CommonPrograms},
	{MENU_FAVORITES,"Menu.Favorites",L"F&avorites",322,MENU_NO,&FOLDERID_Favorites},
	{MENU_DOCUMENTS,"Menu.Documents",L"&Documents",327,MENU_USERFILES,&FOLDERID_Recent},
	{MENU_SETTINGS,"Menu.Settings",L"&Settings",330,MENU_CONTROLPANEL},
	{MENU_SEARCH,"Menu.Search",L"Sear&ch",323,MENU_SEARCH_FILES},
	{MENU_HELP,"Menu.Help",L"&Help and Support",324},
	{MENU_RUN,"Menu.Run",L"&Run...",328},
	{MENU_SEPARATOR},
	{MENU_LOGOFF,"Menu.Logoff",L"&Log Off %s...",325},
	{MENU_UNDOCK,"Menu.Undock",L"Undock Comput&er",331},
	{MENU_DISCONNECT,"Menu.Disconnect",L"D&isconnect...",329},
	{MENU_SHUTDOWN,"Menu.Shutdown",L"Sh&ut Down...",329},
	{MENU_LAST},

	// Documents
	{MENU_USERFILES,NULL,NULL,0,MENU_NO,&FOLDERID_UsersFiles,NULL,"Menu.UserFilesTip",L"Contains folders for Documents, Pictures, Music, and other files that belong to you."},
	{MENU_USERDOCUMENTS,NULL,NULL,0,MENU_NO,&FOLDERID_Documents,NULL,"Menu.UserDocumentsTip",L"Contains letters, reports, and other documents and files."},
	{MENU_USERPICTURES,NULL,NULL,0,MENU_NO,&FOLDERID_Pictures,NULL,"Menu.UserPicturesTip",L"Contains digital photos, images, and graphic files."},
	{MENU_SEPARATOR},
	{MENU_LAST},

	// Settings
	{MENU_CONTROLPANEL,"Menu.ControlPanel",L"&Control Panel",137,MENU_NO,&FOLDERID_ControlPanelFolder},
	{MENU_SEPARATOR},
	{MENU_NETWORK,"Menu.Network",L"&Network Connections",257,MENU_NO,&FOLDERID_ConnectionsFolder,NULL,"Menu.NetworkTip",L"Displays existing network connections on this computer and helps you create new ones"},
	{MENU_PRINTERS,"Menu.Printers",L"&Printers",138,MENU_NO,&FOLDERID_PrintersFolder,NULL,"Menu.PrintersTip",L"Add, remove, and configure local and network printers."},
	{MENU_TASKBAR,"Menu.Taskbar",L"&Taskbar and Start Menu",40,MENU_NO,NULL,NULL,"Menu.TaskbarTip",L"Customize the Start Menu and the taskbar, such as the types of items to be displayed and how they should appear."},
	{MENU_FEATURES,"Menu.Features",L"Programs and &Features",271,MENU_NO,&FOLDERID_ChangeRemovePrograms,NULL,"Menu.FeaturesTip",L"Uninstall or change programs on your computer."},
	{MENU_SEPARATOR},
	{MENU_CLASSIC_SETTINGS,"Menu.ClassicSettings",L"Classic Start &Menu",274,MENU_NO,NULL,NULL,"Menu.SettingsTip",L"Settings for Classic Start Menu"},
	{MENU_LAST},

	// Search
	{MENU_SEARCH_FILES,"Menu.SearchFiles",L"For &Files and Folders...",134},
	{MENU_SEARCH_PRINTER,"Menu.SearchPrinter",L"For &Printer",1006},
	{MENU_SEARCH_COMPUTERS,"Menu.SearchComputers",L"For &Computers",135},
	{MENU_SEARCH_PEOPLE,"Menu.SearchPeople",L"For &People...",269},
	{MENU_LAST},
};

// Options for special menu items
enum
{
	MENU_NONE     = 0,
	MENU_ENABLED  = 1, // the item shows in the menu
	MENU_EXPANDED = 2, // the item is expanded
	MENU_PAGER    = 4, // the menu is always in a pager
};

static StdMenuOption g_StdOptions[]=
{
	{MENU_FAVORITES,MENU_NONE}, // MENU_ENABLED|MENU_EXPANDED from settings, check policy
	{MENU_DOCUMENTS,MENU_NONE}, // MENU_ENABLED|MENU_EXPANDED from settings, check policy
	{MENU_HELP,MENU_ENABLED}, // check policy
	{MENU_RUN,MENU_ENABLED}, // check policy
	{MENU_LOGOFF,MENU_NONE}, // MENU_ENABLED from settings, check policy
	{MENU_DISCONNECT,MENU_NONE}, // MENU_ENABLED if in a remote session, check policy
	{MENU_SHUTDOWN,MENU_ENABLED}, // MENU_NONE if in a remote session, check policy
	{MENU_UNDOCK,MENU_ENABLED}, // from settings, check policy
	{MENU_CONTROLPANEL,MENU_ENABLED|MENU_EXPANDED|MENU_PAGER}, // MENU_EXPANDED from settings, check policy
	{MENU_NETWORK,MENU_ENABLED|MENU_PAGER}, // MENU_EXPANDED from settings, check policy
	{MENU_PRINTERS,MENU_ENABLED|MENU_PAGER}, // MENU_EXPANDED from settings, check policy
	{MENU_TASKBAR,MENU_ENABLED}, // check policy
	{MENU_FEATURES,MENU_ENABLED}, // no setting (prevents the Programs and Features from expanding), check policy (for control panel)
	{MENU_SEARCH,MENU_ENABLED}, // check policy
	{MENU_SEARCH_PRINTER,MENU_NONE}, // MENU_ENABLED if Active Directory is available
	{MENU_SEARCH_COMPUTERS,MENU_NONE}, // MENU_ENABLED if Active Directory is available, check policy
	{MENU_SEARCH_PEOPLE,MENU_NONE}, // MENU_ENABLED if %ProgramFiles%\Windows Mail\wab.exe exists
	{MENU_USERFILES,MENU_ENABLED}, // check policy
	{MENU_USERDOCUMENTS,MENU_ENABLED}, // check policy
	{MENU_USERPICTURES,MENU_ENABLED}, // check policy
};

static int FindStdMenuItem( TMenuID id )
{
	ATLASSERT(id!=MENU_NO && id!=MENU_SEPARATOR && id!=MENU_EMPTY);
	for (int i=0;i<_countof(g_StdMenu);i++)
		if (g_StdMenu[i].id==id)
			return i;
	ATLASSERT(0);
	return -1;
}

// Recent document item (sorts by time, newer first)
struct Document
{
	CString name;
	FILETIME time;

	bool operator<( const Document &x ) { return CompareFileTime(&time,&x.time)>0; }
};

///////////////////////////////////////////////////////////////////////////////

int CMenuContainer::s_MaxRecentDocuments=15;
bool CMenuContainer::s_bScrollMenus=false;
bool CMenuContainer::s_bRTL=false;
bool CMenuContainer::s_bKeyboardCues=false;
bool CMenuContainer::s_bExpandRight=true;
int CMenuContainer::s_MenuBorder=0;
int CMenuContainer::s_MenuStyle=0;
bool CMenuContainer::s_bBehindTaskbar=true;
bool CMenuContainer::s_bShowTopEmpty=false;
bool CMenuContainer::s_bNoEditMenu=false;
bool CMenuContainer::s_bExpandLinks=false;
char CMenuContainer::s_bActiveDirectory=-1;
HTHEME CMenuContainer::s_ThemeMenu;
HTHEME CMenuContainer::s_ThemeList;
COLORREF CMenuContainer::s_MenuColor;
COLORREF CMenuContainer::s_MenuTextColor;
COLORREF CMenuContainer::s_MenuTextHotColor;
COLORREF CMenuContainer::s_MenuTextDColor;
COLORREF CMenuContainer::s_MenuTextHotDColor;
CMenuContainer *CMenuContainer::s_pDragSource=NULL;
std::vector<CMenuContainer*> CMenuContainer::s_Menus;
CComPtr<IShellFolder> CMenuContainer::s_pDesktop;
RECT CMenuContainer::s_MainRect;
DWORD CMenuContainer::s_HoverTime;
CLIPFORMAT CMenuContainer::s_ShellFormat;

// Subclass proc for the toolbars
LRESULT CALLBACK CMenuContainer::ToolbarSubclassProc( HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData )
{
	if (uMsg==WM_MOUSEACTIVATE)
		return MA_NOACTIVATE; // prevent activation with the mouse
	if (uMsg==WM_SYSKEYDOWN)
	{
		// forward system keys to the parent so it can show the keyboard cues
		// WM_KEYDOWN doesn't need to be forwarded because the toolbar control does it through NM_KEYDOWN
		::PostMessage(::GetParent(hWnd),WM_SYSCOMMAND,SC_KEYMENU,0);
	}
	if (uMsg==WM_MOUSEMOVE)
	{
		// ignore mouse messages over separators
		// stops the menu from thinking the mouse is outside of the menu when it is over a separator
		// (but not when the mouse is captured - it interferes with drag/drop)
		POINT pt={(short)LOWORD(lParam),(short)HIWORD(lParam)};
		int idx=(int)::SendMessage(hWnd,TB_HITTEST,0,(LPARAM)&pt);
		if (idx<0 && GetCapture()!=hWnd) return 0;
	}
	return DefSubclassProc(hWnd,uMsg,wParam,lParam);
}

CMenuContainer::CMenuContainer( CMenuContainer *pParent, int options, TMenuID menuID, PIDLIST_ABSOLUTE path1, PIDLIST_ABSOLUTE path2, const CString &regName )
{
	m_RefCount=1;
	m_pParent=pParent;
	m_Options=options;
	m_MenuID=menuID;
	m_RegName=regName;
	m_Bitmap=NULL;
	m_BmpWidth=0;
	m_Path1=ILCloneFull(path1);
	m_Path2=ILCloneFull(path2);

	ATLASSERT(path1 || !path2);
	if (!s_pDesktop)
		SHGetDesktopFolder(&s_pDesktop);
	ATLASSERT(s_pDesktop);
	m_bDestroyed=false;
	s_Menus.push_back(this);

	CoCreateInstance(CLSID_DragDropHelper,NULL,CLSCTX_INPROC_SERVER,IID_IDropTargetHelper,(void**)&m_pDropTargetHelper);
}

CMenuContainer::~CMenuContainer( void )
{
	for (std::vector<MenuItem>::iterator it=m_Items.begin();it!=m_Items.end();++it)
	{
		if (it->pItem1) ILFree(it->pItem1);
		if (it->pItem2) ILFree(it->pItem2);
	}
	if (m_Path1) ILFree(m_Path1);
	if (m_Path2) ILFree(m_Path2);
	if (std::find(s_Menus.begin(),s_Menus.end(),m_pParent)!=s_Menus.end()) // check if the parent is still alive
		m_pParent->m_Submenu=-1;
	if (m_Bitmap) DeleteObject(m_Bitmap);

	// must be here and not in OnDestroy because during drag/drop a menu can close while still processing messages
	s_Menus.erase(std::find(s_Menus.begin(),s_Menus.end(),this));
}

// Initialize the m_Items list
void CMenuContainer::InitItems( void )
{
	m_Items.clear();
	m_bRefreshPosted=false;
	m_Submenu=-1;
	m_HotPos=0;
	m_ContextItem=-1;

	if (m_Options&CONTAINER_DOCUMENTS) // create the recent documents list
	{
		ATLASSERT(m_Path1 && !m_Path2);

		// find all documents

		// with many recent files it takes a long time to go through the IShellFolder enumeration
		// so use FindFirstFile directly
		wchar_t recentPath[_MAX_PATH];
		SHGetPathFromIDList(m_Path1,recentPath);
		wchar_t find[_MAX_PATH];
		swprintf_s(find,L"%s\\*.lnk",recentPath);

		std::vector<Document> docs;

		WIN32_FIND_DATA data;
		HANDLE h=FindFirstFile(find,&data);
		while (h!=INVALID_HANDLE_VALUE)
		{
			Document doc;
			doc.name.Format(L"%s\\%s",recentPath,data.cFileName);
			doc.time=data.ftLastWriteTime;
			docs.push_back(doc);
			if (!FindNextFile(h,&data)) break;
		}

		// sort by time
		std::sort(docs.begin(),docs.end());

		CComPtr<IShellFolder> pFolder;
		s_pDesktop->BindToObject(m_Path1,NULL,IID_IShellFolder,(void**)&pFolder);

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
	if (m_Path1 && !(m_Options&CONTAINER_DOCUMENTS))
	{
		CComPtr<IShellFolder> pFolder;
		s_pDesktop->BindToObject(m_Path1,NULL,IID_IShellFolder,(void**)&pFolder);
		m_pDropFolder=pFolder;

		CComPtr<IEnumIDList> pEnum;
		if (pFolder->EnumObjects(NULL,SHCONTF_NONFOLDERS|SHCONTF_FOLDERS,&pEnum)!=S_OK) pEnum=NULL;

		PITEMID_CHILD pidl;
		while (pEnum && pEnum->Next(1,&pidl,NULL)==S_OK)
		{
			STRRET str;
			if (SUCCEEDED(pFolder->GetDisplayNameOf(pidl,SHGDN_INFOLDER|SHGDN_FORPARSING,&str)))
			{
				wchar_t *name;
				StrRetToStr(&str,pidl,&name);

				MenuItem item={MENU_NO};
				item.icon=g_IconManager.GetIcon(pFolder,pidl,(m_Options&CONTAINER_LARGE)!=0);
				item.pItem1=ILCombine(m_Path1,pidl);

				if (SUCCEEDED(pFolder->GetDisplayNameOf(pidl,SHGDN_INFOLDER|SHGDN_NORMAL,&str)))
				{
					CharUpper(name);
					item.nameHash=CalcFNVHash(name);
					CoTaskMemFree(name);
					StrRetToStr(&str,pidl,&name);
					item.name=name;
					CoTaskMemFree(name);
				}
				else
				{
					item.name=name;
					CharUpper(name);
					item.nameHash=CalcFNVHash(name);
					CoTaskMemFree(name);
				}

				SFGAOF flags=SFGAO_FOLDER|SFGAO_STREAM|SFGAO_LINK; // check if the item is a folder, archive or a link
				if (FAILED(pFolder->GetAttributesOf(1,&pidl,&flags)))
					flags=0;
				item.bLink=(flags&SFGAO_LINK)!=0;
				item.bFolder=(!(m_Options&CONTAINER_NOSUBFOLDERS) && (flags&SFGAO_FOLDER) && (!(flags&(SFGAO_STREAM|SFGAO_LINK)) || (s_bExpandLinks && item.bLink)));
				item.bPrograms=((m_Options&CONTAINER_PROGRAMS) || (!m_pParent && item.id==MENU_NO));

				m_Items.push_back(item);
			}
			ILFree(pidl);
		}
	}

	// add second folder
	if (m_Path2/* && !SHRestricted(REST_NOCOMMONGROUPS)*/)
	{
		CComPtr<IShellFolder> pFolder;
		s_pDesktop->BindToObject(m_Path2,NULL,IID_IShellFolder,(void**)&pFolder);

		CComPtr<IEnumIDList> pEnum;
		if (pFolder->EnumObjects(NULL,SHCONTF_NONFOLDERS|SHCONTF_FOLDERS,&pEnum)!=S_OK) pEnum=NULL;

		PITEMID_CHILD pidl;
		while (pEnum && pEnum->Next(1,&pidl,NULL)==S_OK)
		{
			STRRET str;
			if (SUCCEEDED(pFolder->GetDisplayNameOf(pidl,SHGDN_INFOLDER|SHGDN_FORPARSING,&str)))
			{
				wchar_t *name;
				StrRetToStr(&str,pidl,&name);
				CharUpper(name);
				unsigned int hash=CalcFNVHash(name);
				CoTaskMemFree(name);
				PIDLIST_ABSOLUTE pItem2=ILCombine(m_Path2,pidl);

				// look for another item with the same name
				bool bFound=false;
				for (std::vector<MenuItem>::iterator it=m_Items.begin();it!=m_Items.end();++it)
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
					STRRET str2;
					if (SUCCEEDED(pFolder->GetDisplayNameOf(pidl,SHGDN_INFOLDER|SHGDN_NORMAL,&str2)))
						StrRetToStr(&str2,pidl,&name);
					else
						StrRetToStr(&str,pidl,&name);

					// new item
					MenuItem item={MENU_NO};
					item.icon=g_IconManager.GetIcon(pFolder,pidl,(m_Options&CONTAINER_LARGE)!=0);
					item.name=name;
					item.nameHash=hash;
					item.pItem1=pItem2;

					SFGAOF flags=SFGAO_FOLDER|SFGAO_STREAM|SFGAO_LINK;
					if (FAILED(pFolder->GetAttributesOf(1,&pidl,&flags)))
						flags=0;
					item.bLink=(flags&SFGAO_LINK)!=0;
					item.bFolder=(!(m_Options&CONTAINER_NOSUBFOLDERS) && (flags&SFGAO_FOLDER) && (!(flags&(SFGAO_STREAM|SFGAO_LINK)) || (s_bExpandLinks && item.bLink)));
					item.bPrograms=((m_Options&CONTAINER_PROGRAMS) || (!m_pParent && item.id==MENU_NO));

					m_Items.push_back(item);
					CoTaskMemFree(name);
				}
			}
			ILFree(pidl);
		}
	}

	if (m_Options&CONTAINER_NOPROGRAMS)
	{
		// remove the Programs subfolder from the main menu. it will be added separately
		PIDLIST_ABSOLUTE progs;
		SHGetKnownFolderIDList(FOLDERID_Programs,0,NULL,&progs);
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

	if (m_Items.empty() && m_Path1)
	{
		// add (Empty) item to the empty submenus
		MenuItem item={MENU_EMPTY};
		item.icon=I_IMAGENONE;
		item.name=FindSetting("Menu.Empty",L"(Empty)");
		m_Items.push_back(item);
	}

	// add standard items
	if (m_MenuID!=MENU_NO)
	{
		if (!m_Items.empty())
		{
			MenuItem item={MENU_SEPARATOR};
			m_Items.push_back(item);
		}
		size_t menuIdx=(m_Options&CONTAINER_ADDTOP)?0:m_Items.size();
		for (int idx=FindStdMenuItem(m_MenuID);idx>=0 && g_StdMenu[idx].id!=MENU_LAST;idx++)
		{
			int stdOptions=MENU_ENABLED|MENU_EXPANDED;
			for (int i=0;i<_countof(g_StdOptions);i++)
				if (g_StdOptions[i].id==g_StdMenu[idx].id)
				{
					stdOptions=g_StdOptions[i].options;
					break;
				}

			if (!(stdOptions&MENU_ENABLED)) continue;

			MenuItem item={g_StdMenu[idx].id};
			if (g_StdMenu[idx].icon)
				item.icon=g_IconManager.GetStdIcon(g_StdMenu[idx].icon,(m_Options&CONTAINER_LARGE)!=0);
			else
				item.icon=I_IMAGENONE;

			if (item.id==MENU_LOGOFF)
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
				item.name.Format(FindSetting(g_StdMenu[idx].key,g_StdMenu[idx].name2),user);
			}
			else
				item.name=FindSetting(g_StdMenu[idx].key,g_StdMenu[idx].name2);

			ATLASSERT(g_StdMenu[idx].folder1 || !g_StdMenu[idx].folder2);

			if (g_StdMenu[idx].folder1)
			{
				SHGetKnownFolderIDList(*g_StdMenu[idx].folder1,0,NULL,&item.pItem1);
				if (!g_StdMenu[idx].name2)
				{
					// this is used for the special items in the Documents menu - user docs, user pictures, etc
					// find the real name and icon
					SHFILEINFO info;
					SHGetFileInfo((LPCWSTR)item.pItem1,0,&info,sizeof(info),SHGFI_PIDL|SHGFI_DISPLAYNAME);
					item.name=info.szDisplayName;
					SHGetFileInfo((LPCWSTR)item.pItem1,0,&info,sizeof(info),SHGFI_PIDL|SHGFI_ICONLOCATION);
					item.icon=g_IconManager.GetIcon(info.szDisplayName,info.iIcon,(m_Options&CONTAINER_LARGE)!=0);
				}
			}
			if (g_StdMenu[idx].folder2)
				SHGetKnownFolderIDList(*g_StdMenu[idx].folder2,0,NULL,&item.pItem2);

			item.bFolder=(!(m_Options&CONTAINER_NOSUBFOLDERS) && ((g_StdMenu[idx].folder1 && (stdOptions&MENU_EXPANDED)) || g_StdMenu[idx].submenu!=MENU_NO));
			item.bPrograms=(item.id==MENU_PROGRAMS || item.id==MENU_FAVORITES);
			m_Items.insert(m_Items.begin()+menuIdx,1,item);
			menuIdx++;
		}
	}

	// remove trailing separators
	while (!m_Items.empty() && m_Items[m_Items.size()-1].id==MENU_SEPARATOR)
		m_Items.pop_back();

	// find accelerators
	for (std::vector<MenuItem>::iterator it=m_Items.begin();it!=m_Items.end();++it)
	{
		if (it->id==MENU_SEPARATOR || it->id==MENU_EMPTY || it->name.IsEmpty())
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

// Creates the bitmap to show on the main menu in "large icons" mode
void CMenuContainer::CreateBitmap( int height )
{
	// get the text from the ini file or from the registry
	CRegKey regTitle;
	wchar_t title[256]=L"Windows";
	const wchar_t *setting=FindSetting("Menu.Title",NULL);
	if (setting && *setting)
		wcscpy_s(title,setting);
	else
	{
		ULONG size=_countof(title);
		if (regTitle.Open(HKEY_LOCAL_MACHINE,L"Software\\Microsoft\\Windows NT\\CurrentVersion",KEY_READ)==ERROR_SUCCESS)
			regTitle.QueryStringValue(L"ProductName",title,&size);
	}

	HDC hdc=CreateCompatibleDC(NULL);
	int fontSize=CIconManager::SMALL_ICON_SIZE+8; // font size depending on the DPI

	HFONT font=CreateFont(fontSize,0,0,0,FW_NORMAL,FALSE,FALSE,FALSE,DEFAULT_CHARSET,OUT_DEFAULT_PRECIS,CLIP_DEFAULT_PRECIS,CLEARTYPE_QUALITY,DEFAULT_PITCH,L"Segoe UI");

	HFONT font0=(HFONT)SelectObject(hdc,font);

	RECT rc={0,0,0,0};
	DrawText(hdc,title,-1,&rc,DT_NOPREFIX|DT_SINGLELINE|DT_CALCRECT);
	int width=rc.right;
	rc.right=height;
	rc.bottom+=s_MenuBorder+1; // 1 extra pixel between the title and the menu items

	BITMAPINFO dib={sizeof(dib)};
	dib.bmiHeader.biWidth=rc.right;
	dib.bmiHeader.biHeight=-rc.bottom; // negative height to make it top-down
	dib.bmiHeader.biPlanes=1;
	dib.bmiHeader.biBitCount=32;
	dib.bmiHeader.biCompression=BI_RGB;

	unsigned int *bits;
	HBITMAP bmp=CreateDIBSection(hdc,&dib,DIB_RGB_COLORS,(void**)&bits,NULL,0);

	HBITMAP bmp0=(HBITMAP)SelectObject(hdc,bmp);
	FillRect(hdc,&rc,(HBRUSH)(COLOR_MENU+1));
	int border=s_bRTL?s_MenuBorder:0;

	// create the gradient background
	COLORREF color=GetSysColor(COLOR_MENUHILIGHT);
	int r=color&255, g=(color>>8)&255, b=(color>>16)&255;
	TRIVERTEX verts[]=
	{
		{s_MenuBorder,s_MenuBorder-border,r*64,g*64,b*128,255*256},
		{(width+rc.bottom>s_MenuBorder)?width+rc.bottom:s_MenuBorder,rc.bottom-1-border,r*256,g*256,b*256,255*256},
		{rc.right-s_MenuBorder,s_MenuBorder-border,r*256,g*256,b*256,255*256},
	};
	GRADIENT_RECT rect[2]={{0,1},{1,2}};
	GradientFill(hdc,verts,3,rect,2,GRADIENT_FILL_RECT_H);

	// draw text
	SetTextColor(hdc,0xFFFFFF);
	SetBkMode(hdc,TRANSPARENT);
	rc.left=rc.bottom/2;
	rc.bottom-=border+1;
	rc.top+=s_MenuBorder-border;
	DrawText(hdc,title,-1,&rc,DT_NOPREFIX|DT_VCENTER|DT_SINGLELINE);
	rc.bottom+=border+1;
	rc.top-=s_MenuBorder-border;

	SelectObject(hdc,font0);
	SelectObject(hdc,bmp0);
	DeleteDC(hdc);
	DeleteObject(font);

	// rotate the bitmap 90 degrees. it is possible to draw the text vertically using a rotated font,
	// but unfortunately the text ends up aliased. this way we get better quality. weird, I know.
	unsigned int *bits2;
	dib.bmiHeader.biWidth=rc.bottom;
	dib.bmiHeader.biHeight=-rc.right;
	m_Bitmap=CreateDIBSection(hdc,&dib,DIB_RGB_COLORS,(void**)&bits2,NULL,0);

	for (int y=0;y<rc.right;y++)
		for (int x=0;x<rc.bottom;x++)
			bits2[y*rc.bottom+x]=bits[x*rc.right+rc.right-y-1];

	DeleteObject(bmp);

	m_BmpWidth=rc.bottom;
}

CWindow CMenuContainer::CreateToolbar( int id )
{
	CWindow toolbar=CreateWindowEx(WS_EX_TOOLWINDOW,TOOLBARCLASSNAME,L"",WS_CHILD|WS_VISIBLE|TBSTYLE_TOOLTIPS|TBSTYLE_FLAT|TBSTYLE_LIST|TBSTYLE_TRANSPARENT|TBSTYLE_REGISTERDROP|CCS_NODIVIDER|CCS_NOPARENTALIGN|CCS_NORESIZE|CCS_VERT,0,0,10,10,m_hWnd,(HMENU)id,g_Instance,NULL);
	SetWindowTheme(toolbar.m_hWnd,L"",L""); // disable cross-fade on Vista
	toolbar.SendMessage(TB_BUTTONSTRUCTSIZE,sizeof(TBBUTTON),0);

	HIMAGELIST images=(m_Options&CONTAINER_LARGE)?g_IconManager.m_LargeIcons:g_IconManager.m_SmallIcons;
	toolbar.SendMessage(TB_SETIMAGELIST,0,(LPARAM)images);

	SetWindowSubclass(toolbar.m_hWnd,ToolbarSubclassProc,(UINT_PTR)this,0);

	return toolbar;
}

// Create one or more toolbars and the pager. Reuses the existing controls if possible
void CMenuContainer::InitToolbars( void )
{
	// add buttons
	int maxHeight=(s_MainRect.bottom-s_MainRect.top);
	CWindow toolbar;
	std::vector<CWindow> oldToolbars;
	oldToolbars.swap(m_Toolbars);
	int column=0;
	for (size_t i=0;i<m_Items.size();i++)
	{
		if (!toolbar)
		{
			if (m_Items[i].id==MENU_SEPARATOR)
				continue; // don't start with a separator
			if (oldToolbars.empty())
			{
				toolbar=CreateToolbar(column+ID_TOOLBAR_FIRST);
			}
			else
			{
				// reuse existing toolbar
				toolbar=oldToolbars[0];
				toolbar.SetParent(m_hWnd);
				oldToolbars.erase(oldToolbars.begin());
				for (int j=(int)toolbar.SendMessage(TB_BUTTONCOUNT)-1;j>=0;j--)
					toolbar.SendMessage(TB_DELETEBUTTON,j);
			}
			if (!s_bKeyboardCues)
				toolbar.SendMessage(TB_SETDRAWTEXTFLAGS,DT_HIDEPREFIX,DT_HIDEPREFIX);
			m_Toolbars.push_back(toolbar);
		}

		// add the new button
		TBBUTTON button={m_Items[i].icon,(int)i+ID_OFFSET,TBSTATE_ENABLED|TBSTATE_WRAP,BTNS_BUTTON,{0},i+ID_OFFSET,(INT_PTR)(const wchar_t*)m_Items[i].name};
		if (m_Items[i].id==MENU_SEPARATOR)
			button.fsStyle=BTNS_SEP;
		else if (m_Items[i].id==MENU_NO)
			button.fsStyle|=BTNS_NOPREFIX;
		if (!s_bShowTopEmpty && !m_pParent && (m_Items[i].id==MENU_EMPTY || (i==1 && m_Items[0].id==MENU_EMPTY)))
			button.fsState|=TBSTATE_HIDDEN; // this is the first (Empty) item in the top menu. hide it for now
		int n=(int)toolbar.SendMessage(TB_BUTTONCOUNT);
		m_Items[i].column=column;
		m_Items[i].btnIndex=n;
		toolbar.SendMessage(TB_INSERTBUTTON,n,(LPARAM)&button);

		if (!(m_Options&CONTAINER_PAGER))
		{
			RECT rc;
			toolbar.SendMessage(TB_GETITEMRECT,n,(LPARAM)&rc);
			if (rc.bottom>maxHeight && n>0)
			{
				// too tall, delete the last button and start a new column
				toolbar.SendMessage(TB_DELETEBUTTON,n);
				i--;
				column++;
				toolbar=NULL;
			}
		}
	}

	// destroy the unused toolbars
	for (std::vector<CWindow>::iterator it=oldToolbars.begin();it!=oldToolbars.end();++it)
		it->DestroyWindow();
	oldToolbars.clear();

	// calculate toolbar size
	int maxw=0, maxh=0, maxbh=0;
	for (std::vector<CWindow>::iterator it=m_Toolbars.begin();it!=m_Toolbars.end();++it)
	{
		it->SendMessage(TB_AUTOSIZE);
		LRESULT n=it->SendMessage(TB_BUTTONCOUNT);
		RECT rc;
		it->SendMessage(TB_GETITEMRECT,n-1,(LPARAM)&rc);
		if (maxw<rc.right) maxw=rc.right; // button width
		if (maxh<rc.bottom) maxh=rc.bottom; // toolbar height
		if (maxbh<rc.bottom-rc.top) maxbh=rc.bottom-rc.top; // button height
	}

	maxw+=16; // add 16 pixels for the submenu arrow

	if (m_Bitmap)
	{
		DeleteObject(m_Bitmap);
		m_Bitmap=NULL;
	}

	m_BmpWidth=0;
	if ((m_Options&CONTAINER_LARGE) && !m_pParent)
		CreateBitmap((maxh<maxHeight?maxh:maxHeight)+2*s_MenuBorder);

	int offs=0, step=maxw;
	if (m_BmpWidth)
		offs=m_BmpWidth;
	else
		offs=s_MenuBorder;

	int offs0=offs;

	for (std::vector<CWindow>::iterator it=m_Toolbars.begin();it!=m_Toolbars.end();++it, offs+=step)
	{
		it->SendMessage(TB_SETBUTTONSIZE,0,MAKELONG(maxw,maxbh));
		it->SetWindowPos(NULL,offs,s_MenuBorder,maxw,maxh,SWP_NOZORDER|SWP_NOACTIVATE);
	}

	// create pager
	if ((m_Options&CONTAINER_PAGER) && maxh>maxHeight)
	{
		if (!m_Pager)
			m_Pager=CreateWindow(WC_PAGESCROLLER,L"",WS_CHILD|WS_VISIBLE|PGS_DRAGNDROP|PGS_AUTOSCROLL|PGS_VERT|CCS_NODIVIDER|CCS_NOPARENTALIGN|CCS_NORESIZE,offs0,s_MenuBorder,maxw,maxHeight,m_hWnd,(HMENU)ID_PAGER,g_Instance,NULL);
		m_Pager.SendMessage(PGM_SETBUTTONSIZE,0,maxbh/2);
		m_Toolbars[0].SetParent(m_Pager);
		m_Toolbars[0].SetWindowPos(NULL,0,0,0,0,SWP_NOZORDER|SWP_NOSIZE);
		m_Pager.SendMessage(PGM_SETCHILD,0,(LPARAM)m_Toolbars[0].m_hWnd);
		maxh=maxHeight;
	}
	else if (m_Pager.m_hWnd)
	{
		m_Pager.DestroyWindow();
		m_Pager=NULL;
	}

	RECT rc={0,0,maxw*(int)m_Toolbars.size()+offs0+s_MenuBorder,maxh+s_MenuBorder*2};
	AdjustWindowRectEx(&rc,s_MenuStyle,FALSE,WS_EX_TOOLWINDOW|WS_EX_TOPMOST);
	RECT rc0;
	GetWindowRect(&rc0);
	OffsetRect(&rc,(m_Options&CONTAINER_LEFT)?(rc0.left-rc.left):(rc0.right-rc.right),(m_Options&CONTAINER_TOP)?(rc0.top-rc.top):(rc0.bottom-rc.bottom));
	SetWindowPos(NULL,&rc,SWP_NOZORDER|SWP_NOACTIVATE|SWP_NOCOPYBITS);
}

LRESULT CMenuContainer::OnCreate( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled )
{
	if (!m_pParent)
	{
		s_MenuColor=GetSysColor(COLOR_MENU);
		s_MenuTextColor=GetSysColor(COLOR_MENUTEXT);
		s_MenuTextHotColor=GetSysColor(COLOR_HIGHLIGHTTEXT);
		s_MenuTextDColor=GetSysColor(COLOR_GRAYTEXT);
		s_MenuTextHotDColor=GetSysColor(COLOR_HIGHLIGHTTEXT);
		if (m_Options&CONTAINER_THEME)
		{
			SetWindowTheme(m_hWnd,L"explorer",NULL); // get the nice "explorer" theme for the listview selection

			HTHEME s_ThemeMenu=OpenThemeData(m_hWnd,L"menu");
			if (s_ThemeMenu)
			{
				s_MenuColor=GetThemeSysColor(s_ThemeMenu,COLOR_MENU);
				s_MenuTextColor=GetThemeSysColor(s_ThemeMenu,COLOR_MENUTEXT);
				s_MenuTextHotColor=s_MenuTextColor;
				s_MenuTextDColor=GetThemeSysColor(s_ThemeMenu,COLOR_GRAYTEXT);
				s_MenuTextHotDColor=s_MenuTextDColor;
				CloseThemeData(s_ThemeMenu);
			}

			s_ThemeList=OpenThemeData(m_hWnd,L"listview");
		}
	}
	InitToolbars();
	m_ClickTime=GetMessageTime()-10000;
	m_ClickPos.x=m_ClickPos.y=-20;
	m_HotPos=-1;
	return 0;
}

LRESULT CMenuContainer::OnCustomDraw( int idCtrl, LPNMHDR pnmh, BOOL& bHandled )
{
	LRESULT res=CDRF_DODEFAULT;
	bHandled=FALSE;
	NMTBCUSTOMDRAW *pDraw=(NMTBCUSTOMDRAW*)pnmh;
	if (pDraw->nmcd.dwDrawStage==CDDS_PREPAINT)
	{
		res=CDRF_NOTIFYITEMDRAW;
		bHandled=TRUE;
	}
	else if (pDraw->nmcd.dwDrawStage==CDDS_ITEMPREPAINT)
	{
		res=TBCDRF_USECDCOLORS;
		bool bDisabled=(pDraw->nmcd.uItemState&CDIS_DISABLED) || (pDraw->nmcd.lItemlParam>=ID_OFFSET && m_Items[pDraw->nmcd.lItemlParam-ID_OFFSET].id==MENU_EMPTY);
		if (pDraw->nmcd.uItemState&CDIS_HOT)
		{
			pDraw->clrText=bDisabled?s_MenuTextHotDColor:s_MenuTextHotColor;
			if (s_ThemeList)
			{
				// draw the background for the hot item
				DrawThemeBackground(s_ThemeList,pDraw->nmcd.hdc,LVP_LISTITEM,LISS_HOTSELECTED,&pDraw->nmcd.rc,NULL);
			}
			else
			{
				// set the color for the hot item
				pDraw->clrHighlightHotTrack=GetSysColor(COLOR_HIGHLIGHT);
				res|=TBCDRF_HILITEHOTTRACK;
			}
		}
		else
		{
			// set the color for the (Empty) items
			pDraw->clrText=bDisabled?s_MenuTextDColor:s_MenuTextColor;
		}
		res|=CDRF_NOTIFYPOSTPAINT|TBCDRF_NOMARK|TBCDRF_NOOFFSET|TBCDRF_NOEDGES;
		bHandled=TRUE;
	}
	else if (pDraw->nmcd.dwDrawStage==CDDS_ITEMPOSTPAINT && pDraw->nmcd.lItemlParam>=ID_OFFSET && m_Items[pDraw->nmcd.lItemlParam-ID_OFFSET].bFolder)
	{
		// draw a small triangle for the submenus
		bool bDisabled=(pDraw->nmcd.uItemState&CDIS_DISABLED) || (pDraw->nmcd.lItemlParam>=ID_OFFSET && m_Items[pDraw->nmcd.lItemlParam-ID_OFFSET].id==MENU_EMPTY);
		if (s_ThemeMenu)
		{
			RECT rc=pDraw->nmcd.rc;
			rc.left=rc.right-16;
			DrawThemeBackground(s_ThemeMenu,pDraw->nmcd.hdc,MENU_POPUPSUBMENU,MSM_NORMAL,&rc,NULL);
		}
		else
		{
			if (pDraw->nmcd.uItemState&CDIS_HOT)
				SetDCBrushColor(pDraw->nmcd.hdc,bDisabled?s_MenuTextHotDColor:s_MenuTextHotColor);
			else
				SetDCBrushColor(pDraw->nmcd.hdc,bDisabled?s_MenuTextDColor:s_MenuTextColor);
			SelectObject(pDraw->nmcd.hdc,GetStockObject(DC_BRUSH));
			SelectObject(pDraw->nmcd.hdc,GetStockObject(NULL_PEN));
			int x=pDraw->nmcd.rc.right-8;
			int y=(pDraw->nmcd.rc.top+pDraw->nmcd.rc.bottom)/2;
			POINT points[3];
			points[0].x=x;
			points[0].y=y-4;
			points[1].x=x;
			points[1].y=y+4;
			points[2].x=x+4;
			points[2].y=y;
			Polygon(pDraw->nmcd.hdc,points,3);
		}
		bHandled=TRUE;
	}
	return res;
}

static CString g_RenameText;
static const POINT *g_RenamePos;

// Dialog proc for the Rename dialog box
static INT_PTR CALLBACK RenameDlgProc( HWND hwndDlg, UINT uMsg, WPARAM wParam, LPARAM lParam )
{
	if (uMsg==WM_INITDIALOG)
	{
		// translate text
		SetWindowText(hwndDlg,FindSetting("Menu.RenameTitle",L"Rename"));
		SetDlgItemText(hwndDlg,IDC_LABEL,FindSetting("Menu.RenamePrompt",L"&New name:"));
		SetDlgItemText(hwndDlg,IDOK,FindSetting("Menu.RenameOK",L"OK"));
		SetDlgItemText(hwndDlg,IDCANCEL,FindSetting("Menu.RenameCancel",L"Cancel"));
		SetDlgItemText(hwndDlg,IDC_EDITNAME,g_RenameText);
		if (g_RenamePos)
		{
			// position near the item
			SetWindowPos(hwndDlg,NULL,g_RenamePos->x,g_RenamePos->y,0,0,SWP_NOZORDER|SWP_NOSIZE);
			SendMessage(hwndDlg,DM_REPOSITION,0,0);
		}
		return TRUE;
	}
	if (uMsg==WM_COMMAND && wParam==IDOK)
	{
		wchar_t buf[1024];
		GetDlgItemText(hwndDlg,IDC_EDITNAME,buf,_countof(buf));
		g_RenameText=buf;

		EndDialog(hwndDlg,1);
		return TRUE;
	}
	if (uMsg==WM_COMMAND && wParam==IDCANCEL)
	{
		EndDialog(hwndDlg,0);
		return TRUE;
	}
	return FALSE;
}

// Dialog proc for the Log Off dialog box
static INT_PTR CALLBACK LogOffDlgProc( HWND hwndDlg, UINT uMsg, WPARAM wParam, LPARAM lParam )
{
	if (uMsg==WM_INITDIALOG)
	{
		// translate text
		SendDlgItemMessage(hwndDlg,IDC_STATICICON,STM_SETICON,lParam,0);
		SetWindowText(hwndDlg,FindSetting("Menu.LogoffTitle",L"Log Off Windows"));
		SetDlgItemText(hwndDlg,IDC_PROMPT,FindSetting("Menu.LogoffPrompt",L"Are you sure you want to log off?"));
		SetDlgItemText(hwndDlg,IDOK,FindSetting("Menu.LogoffYes",L"&Log Off"));
		SetDlgItemText(hwndDlg,IDCANCEL,FindSetting("Menu.LogoffNo",L"&No"));
		return TRUE;
	}
	if (uMsg==WM_COMMAND && wParam==IDOK)
	{
		EndDialog(hwndDlg,1);
		return TRUE;
	}
	if (uMsg==WM_COMMAND && wParam==IDCANCEL)
	{
		EndDialog(hwndDlg,0);
		return TRUE;
	}
	return FALSE;
}

// This function "activates" an item. The item can be activated in multiple ways:
// ACTIVATE_SELECT - select the item, make sure it is visible
// ACTIVATE_OPEN - if the item is a submenu, it is opened. otherwise the item is just selected (but all submenus are closed first)
// ACTIVATE_OPEN_KBD - same as above, but when done with a keyboard
// ACTIVATE_EXECUTE - executes the item. it can be a shell item or a command item
// ACTIVATE_MENU - shows the context menu for the item
void CMenuContainer::ActivateItem( int index, TActivateType type, const POINT *pPt )
{
	if (index<0)
	{
		if (type==ACTIVATE_SELECT)
		{
			m_Toolbars[0].SetFocus();
			m_Toolbars[0].SendMessage(TB_SETHOTITEM,-1);
		}
		return;
	}

	MenuItem &item=m_Items[index];
	if (type==ACTIVATE_SELECT)
	{
		// set the hot item
		m_Toolbars[item.column].SetFocus();
		m_Toolbars[item.column].SendMessage(TB_SETHOTITEM,item.btnIndex);
		for (size_t i=0;i<m_Toolbars.size();i++)
			if (i!=item.column)
				m_Toolbars[i].SendMessage(TB_SETHOTITEM,-1);
		if (m_Pager.m_hWnd)
		{
			// scroll the pager to make this item visible
			RECT rc;
			m_Pager.GetClientRect(&rc);
			int height=rc.bottom;
			m_Toolbars[item.column].SendMessage(TB_GETITEMRECT,item.btnIndex,(LPARAM)&rc);
			int btnSize=(int)m_Pager.SendMessage(PGM_GETBUTTONSIZE);
			int pos=(int)m_Pager.SendMessage(PGM_GETPOS);
			if (rc.top<pos)
				m_Pager.SendMessage(PGM_SETPOS,0,rc.top);
			else if (rc.bottom+btnSize>pos+height)
				m_Pager.SendMessage(PGM_SETPOS,0,rc.bottom+btnSize-height);
		}
		return;
	}

	if (type==ACTIVATE_OPEN || type==ACTIVATE_OPEN_KBD)
	{
		m_HotPos=GetMessagePos();
		if (!item.bFolder)
		{
			SetActiveWindow();
			// destroy old submenus
			for (int i=(int)s_Menus.size()-1;s_Menus[i]!=this;i--)
				if (!s_Menus[i]->m_bDestroyed)
					s_Menus[i]->DestroyWindow();

			// just select the item
			ActivateItem(index,ACTIVATE_SELECT,NULL);
			return;
		}

		// open a submenu - create a new menu object
		TMenuID subMenuID=MENU_NO;
		if (item.id!=MENU_NO)
		{
			// expand standard menu
			int stdIndex=FindStdMenuItem(item.id);
			subMenuID=g_StdMenu[stdIndex].submenu;
		}

		int options=0;
		if (item.id==MENU_CONTROLPANEL)
			options|=CONTAINER_NOSUBFOLDERS;
		if (item.id==MENU_DOCUMENTS)
			options|=CONTAINER_DOCUMENTS|CONTAINER_NOSUBFOLDERS|CONTAINER_ADDTOP;
		if (item.bPrograms)
			options|=CONTAINER_PROGRAMS|CONTAINER_DRAG|CONTAINER_DROP;
		if (item.id==MENU_CONTROLPANEL)
			options|=CONTAINER_DRAG;
		if (item.bLink || (m_Options&CONTAINER_LINK))
			options|=CONTAINER_LINK;

		if (s_bScrollMenus || (options&CONTAINER_LINK))
			options|=CONTAINER_PAGER; // scroll all menus
		else
		{
			// scroll only selected menus
			for (int i=0;i<_countof(g_StdOptions);i++)
				if (g_StdOptions[i].id==item.id)
				{
					if (g_StdOptions[i].options&MENU_PAGER)
						options|=CONTAINER_PAGER;
					break;
				}
		}

		CMenuContainer *pMenu=new CMenuContainer(this,options,subMenuID,item.pItem1,item.pItem2,m_RegName+L"\\"+item.name);
		pMenu->InitItems();

		BOOL animate;
		if (m_Submenu>=0 || s_bRTL)
			animate=FALSE;
		else
			SystemParametersInfo(SPI_GETMENUANIMATION,NULL,&animate,0);

		// destroy old submenus
		SetActiveWindow();
		for (int i=(int)s_Menus.size()-2;s_Menus[i]!=this;i--)
			if (!s_Menus[i]->m_bDestroyed)
				s_Menus[i]->DestroyWindow();

		// open submenu
		CWindow toolbar=m_Toolbars[item.column];

		pMenu->Create(NULL,NULL,NULL,s_MenuStyle,WS_EX_TOOLWINDOW|WS_EX_TOPMOST|(s_bRTL?WS_EX_LAYOUTRTL:0));
		RECT rc2;
		pMenu->GetWindowRect(&rc2);

		RECT border={-s_MenuBorder,-s_MenuBorder,s_MenuBorder,s_MenuBorder};
		AdjustWindowRectEx(&border,s_MenuStyle,FALSE,WS_EX_TOOLWINDOW|WS_EX_TOPMOST);

		// position new menu
		int w=rc2.right-rc2.left;
		int h=rc2.bottom-rc2.top;
		DWORD fade;
		SystemParametersInfo(SPI_GETMENUFADE,NULL,&fade,0);
		DWORD animFlags=AW_ACTIVATE|(fade?AW_BLEND:AW_SLIDE);

		RECT rc;
		toolbar.SendMessage(TB_GETITEMRECT,item.btnIndex,(LPARAM)&rc);
		toolbar.ClientToScreen(&rc);
		if (s_bRTL)
		{
			// rc.left and rc.right coming from TB_GETITEMRECT are swapped
			int q=rc.left; rc.left=rc.right; rc.right=q;
		}

		if (s_bExpandRight)
		{
			if (rc.right+border.left+w<=s_MainRect.right)
			{
				// right
				rc2.left=rc.right+border.left;
				rc2.right=rc2.left+w;
				animFlags|=AW_HOR_POSITIVE;
				pMenu->m_Options|=CONTAINER_LEFT;
			}
			else if (rc.left+border.right-w>=s_MainRect.left)
			{
				// left
				rc2.right=rc.left+border.right;
				rc2.left=rc2.right-w;
				animFlags|=AW_HOR_NEGATIVE;
			}
			else
			{
				// right again
				rc2.right=s_MainRect.right;
				rc2.left=rc2.right-w;
				animFlags|=AW_HOR_POSITIVE;
				pMenu->m_Options|=CONTAINER_LEFT;
			}
		}
		else
		{
			if (rc.left+border.right-w>=s_MainRect.left)
			{
				// left
				rc2.right=rc.left+border.right;
				rc2.left=rc2.right-w;
				animFlags|=AW_HOR_NEGATIVE;
			}
			else if (rc.right+border.left+w<=s_MainRect.right)
			{
				// right
				rc2.left=rc.right+border.left;
				rc2.right=rc2.left+w;
				animFlags|=AW_HOR_POSITIVE;
				pMenu->m_Options|=CONTAINER_LEFT;
			}
			else
			{
				// left again
				rc2.left=s_MainRect.left;
				rc2.right=rc2.left+w;
				animFlags|=AW_HOR_NEGATIVE;
			}
		}

		if (s_bRTL)
			animFlags^=(AW_HOR_POSITIVE|AW_HOR_NEGATIVE); // RTL flips the animation

		if (rc.top+border.top+h<=s_MainRect.bottom)
		{
			// down
			rc2.top=rc.top+border.top;
			rc2.bottom=rc2.top+h;
			pMenu->m_Options|=CONTAINER_TOP;
		}
		else if (rc.bottom+border.bottom-h>=s_MainRect.top)
		{
			// up
			rc2.bottom=rc.bottom+border.bottom;
			rc2.top=rc2.bottom-h;
		}
		else
		{
			// down again
			rc2.bottom=s_MainRect.bottom;
			rc2.top=rc2.bottom-h;
			pMenu->m_Options|=CONTAINER_TOP;
		}

		m_Toolbars[item.column].SendMessage(TB_SETHOTITEM,item.btnIndex);
		for (size_t i=0;i<m_Toolbars.size();i++)
		{
			if (i!=item.column)
				m_Toolbars[i].SendMessage(TB_SETHOTITEM,-1);
			m_Toolbars[i].UpdateWindow();
		}

		m_Submenu=index;

		if (animate)
		{
			pMenu->SetWindowPos(HWND_TOPMOST,&rc2,0);
			AnimateWindow(pMenu->m_hWnd,MENU_ANIM_SPEED_SUBMENU,animFlags);
		}
		else
			pMenu->SetWindowPos(HWND_TOPMOST,&rc2,SWP_SHOWWINDOW);
		pMenu->m_Toolbars[0].SetFocus();
		if (type==ACTIVATE_OPEN_KBD)
			pMenu->m_Toolbars[0].SendMessage(TB_SETHOTITEM,0);
		else
			pMenu->m_Toolbars[0].SendMessage(TB_SETHOTITEM,-1);
		return;
	}

	if (type==ACTIVATE_EXECUTE)
	{
		if (item.id==MENU_EMPTY) return;
		// when executing an item close the whole menu
		if (GetKeyState(VK_SHIFT)>=0)
		{
			for (std::vector<CMenuContainer*>::reverse_iterator it=s_Menus.rbegin();it!=s_Menus.rend();++it)
				if (!(*it)->m_bDestroyed)
					(*it)->PostMessage(WM_CLOSE);
		}
	}

	if (type==ACTIVATE_MENU)
	{
		// when showing the context menu close all submenus
		SetActiveWindow();
		for (int i=(int)s_Menus.size()-1;s_Menus[i]!=this;i--)
			if (!s_Menus[i]->m_bDestroyed)
				s_Menus[i]->DestroyWindow();
	}

	m_Toolbars[item.column].SendMessage(TB_SETHOTITEM,item.btnIndex);

	if (!item.pItem1)
	{
		// regular command item
		if (type!=ACTIVATE_EXECUTE) return;

		// special handling for command items
		CComPtr<IShellDispatch2> pShellDisp;
		switch (item.id)
		{
			case MENU_TASKBAR: // show taskbar properties
				if (SUCCEEDED(CoCreateInstance(CLSID_Shell,NULL,CLSCTX_SERVER,IID_IShellDispatch2,(void**)&pShellDisp)))
					pShellDisp->TrayProperties();
				break;
			case MENU_CLASSIC_SETTINGS: // show our settings
				EditSettings();
				break;
			case MENU_SEARCH_FILES: // show the search UI
				{
					const wchar_t *command=FindSetting("Command.SearchFiles",L"");
					if (*command)
					{
						wchar_t buf[1024];
						wcscpy_s(buf,command);
						DoEnvironmentSubst(buf,_countof(buf));
						ShellExecute(NULL,NULL,buf,NULL,NULL,SW_SHOWNORMAL);
					}
					else
					{
						if (SUCCEEDED(CoCreateInstance(CLSID_Shell,NULL,CLSCTX_SERVER,IID_IShellDispatch2,(void**)&pShellDisp)))
							pShellDisp->FindFiles();
					}
				}
				break;
			case MENU_SEARCH_PRINTER: // search for network printers
				if (SUCCEEDED(CoCreateInstance(CLSID_Shell,NULL,CLSCTX_SERVER,IID_IShellDispatch2,(void**)&pShellDisp)))
				pShellDisp->FindPrinter(CComBSTR(L""),CComBSTR(L""),CComBSTR(L""));
				break;
			case MENU_SEARCH_COMPUTERS: // search for computers
				if (SUCCEEDED(CoCreateInstance(CLSID_Shell,NULL,CLSCTX_SERVER,IID_IShellDispatch2,(void**)&pShellDisp)))
					pShellDisp->FindComputer();
				break;
			case MENU_SEARCH_PEOPLE: // search for people using Windows Mail
				{
					SHELLEXECUTEINFO execute={sizeof(execute),SEE_MASK_DOENVSUBST,m_hWnd,L"open"};
					execute.lpFile=L"%ProgramFiles%\\Windows Mail\\wab.exe";
					execute.lpParameters=L"/find";
					execute.lpDirectory=L"%ProgramFiles%\\Windows Mail";
					execute.nShow=SW_SHOWNORMAL;
					ShellExecuteEx(&execute);
				}
				break;
			case MENU_HELP: // show Windows help
				if (SUCCEEDED(CoCreateInstance(CLSID_Shell,NULL,CLSCTX_SERVER,IID_IShellDispatch2,(void**)&pShellDisp)))
					pShellDisp->Help();
				break;
			case MENU_RUN: // show the Run box
				if (SUCCEEDED(CoCreateInstance(CLSID_Shell,NULL,CLSCTX_SERVER,IID_IShellDispatch2,(void**)&pShellDisp)))
					pShellDisp->FileRun();
				break;
			case MENU_LOGOFF: // log off
				{
					if (!(m_Options&CONTAINER_CONFIRM_LO))
						ExitWindowsEx(EWX_LOGOFF,0);
					else
					{
						HMODULE hShell32=GetModuleHandle(L"Shell32.dll");
						HICON icon=LoadIcon(hShell32,MAKEINTRESOURCE(45));
						INT_PTR res=DialogBoxParam(g_Instance,MAKEINTRESOURCE(s_bRTL?IDD_LOGOFFR:IDD_LOGOFF),NULL,LogOffDlgProc,(LPARAM)icon);
						DestroyIcon(icon);
						if (res)
							ExitWindowsEx(EWX_LOGOFF,0);
					}
				}
				break;
			case MENU_DISCONNECT: // disconnect the current Terminal Services session (remote desktop)
				WTSDisconnectSession(WTS_CURRENT_SERVER_HANDLE,WTS_CURRENT_SESSION,FALSE);
				break;
			case MENU_UNDOCK: // undock the PC
				if (SUCCEEDED(CoCreateInstance(CLSID_Shell,NULL,CLSCTX_SERVER,IID_IShellDispatch2,(void**)&pShellDisp)))
					pShellDisp->EjectPC();
				break;
			case MENU_SHUTDOWN: // shutdown - ask to shutdown, log off, sleep, etc
				if (SUCCEEDED(CoCreateInstance(CLSID_Shell,NULL,CLSCTX_SERVER,IID_IShellDispatch2,(void**)&pShellDisp)))
					pShellDisp->ShutdownWindows();
				break;
		}
		return;
	}

	// create a context menu for the selected item. the context menu can be shown (ACTIVATE_MENU) or its default
	// item can be executed automatically (ACTIVATE_EXECUTE)
	CComPtr<IContextMenu> pMenu;
	HMENU menu=CreatePopupMenu();

	CComPtr<IShellFolder> pFolder;
	PCUITEMID_CHILD pidl;

	if (item.bFolder && item.pItem2)
	{
		// context menu for a double folder - show Open, Open All Users, etc.
		AppendMenu(menu,MF_STRING,CMD_OPEN,FindSetting("Menu.Open",L"&Open"));
		AppendMenu(menu,MF_STRING,CMD_EXPLORE,FindSetting("Menu.Explore",L"&Explore"));
		AppendMenu(menu,MF_SEPARATOR,0,0);
		AppendMenu(menu,MF_STRING,CMD_OPEN_ALL,FindSetting("Menu.OpenAll",L"O&pen All Users"));
		AppendMenu(menu,MF_STRING,CMD_EXPLORE_ALL,FindSetting("Menu.ExploreAll",L"E&xplore All Users"));
		AppendMenu(menu,MF_SEPARATOR,0,0);
		AppendMenu(menu,MF_STRING,CMD_PROPERTIES,FindSetting("Menu.Properties",L"P&roperties"));
		SetMenuDefaultItem(menu,CMD_OPEN,FALSE);
	}
	else
	{
		SHBindToParent(item.pItem1,IID_IShellFolder,(void**)&pFolder,&pidl);
		if (FAILED(pFolder->GetUIObjectOf(m_hWnd,1,&pidl,IID_IContextMenu,NULL,(void**)&pMenu)))
		{
			DestroyMenu(menu);
			return;
		}

		UINT flags=CMF_DEFAULTONLY;
		if (type==ACTIVATE_MENU)
		{
			flags=CMF_NORMAL|CMF_CANRENAME;
			if (GetKeyState(VK_SHIFT)<0) flags|=CMF_EXTENDEDVERBS;
		}
		if (FAILED(pMenu->QueryContextMenu(menu,0,CMD_LAST,32767,flags)))
		{
			DestroyMenu(menu);
			return;
		}
	}

	int res=0;
	if (type==ACTIVATE_EXECUTE)
	{
		// just pick the default item
		res=GetMenuDefaultItem(menu,FALSE,0);
		if (res<0) res=0;
	}
	else
	{
		// show the context menu
		m_pMenu2=pMenu;
		m_pMenu3=pMenu;
		if (item.id==MENU_NO) // clicked on a movable item
		{
			int n=0;
			for (std::vector<MenuItem>::const_iterator it=m_Items.begin();it!=m_Items.end();++it)
				if (it->id==MENU_NO)
					n++;
			if (n>1)
			{
				// more than 1 movable items
				AppendMenu(menu,MF_SEPARATOR,0,0);
				AppendMenu(menu,MF_STRING,CMD_SORT,FindSetting("Menu.SortByName",L"Sort &by Name"));
			}
		}
		TPMPARAMS params={sizeof(params)}, *pParams=NULL;
		POINT pt2=*pPt;
		if (pt2.x==-1 && pt2.y==-1)
		{
			CWindow toolbar=m_Toolbars[item.column];
			toolbar.SendMessage(TB_GETITEMRECT,item.btnIndex,(LPARAM)&params.rcExclude);
			toolbar.ClientToScreen(&params.rcExclude);
			pt2.x=params.rcExclude.left;
			pt2.y=params.rcExclude.top;
			pParams=&params;
		}
		m_ContextItem=index;
		KillTimer(TIMER_HOVER);
		res=TrackPopupMenuEx(menu,TPM_LEFTBUTTON|TPM_RETURNCMD|(IsLanguageRTL()?TPM_LAYOUTRTL:0),pt2.x,pt2.y,m_hWnd,pParams);
		m_ContextItem=-1;
		if (m_pMenu2) m_pMenu2.Release();
		if (m_pMenu3) m_pMenu3.Release();
	}

	if (GetKeyState(VK_SHIFT)<0)
		LockSetForegroundWindow(LSFW_LOCK);

	// handle our standard commands
	if (res==CMD_OPEN || res==CMD_OPEN_ALL || res==CMD_EXPLORE || res==CMD_EXPLORE_ALL || res==CMD_PROPERTIES)
	{
		PIDLIST_ABSOLUTE pidl;
		if (res!=CMD_OPEN_ALL && res!=CMD_EXPLORE_ALL)
			pidl=item.pItem1;
		else
			pidl=item.pItem2;

		SHELLEXECUTEINFO execute={sizeof(execute),SEE_MASK_IDLIST|SEE_MASK_INVOKEIDLIST};
		if (res==CMD_OPEN || res==CMD_OPEN_ALL)
			execute.lpVerb=L"open";
		else if (res==CMD_EXPLORE || res==CMD_EXPLORE_ALL)
			execute.lpVerb=L"explore";
		else
			execute.lpVerb=L"properties";
		execute.lpIDList=pidl;
		execute.nShow=SW_SHOWNORMAL;
		ShellExecuteEx(&execute);
	}

	if (res==CMD_SORT)
	{
		std::vector<SortMenuItem> items;
		for (std::vector<MenuItem>::const_iterator it=m_Items.begin();it!=m_Items.end();++it)
			if (it->id==MENU_NO)
			{
				SortMenuItem item={it->name,it->nameHash,it->bFolder};
				items.push_back(item);
			}
		std::sort(items.begin(),items.end());
		SaveItemOrder(items);
		PostRefreshMessage();
	}

	// handle the shell commands
	if (res>=CMD_LAST)
	{
		// handle special verbs
		char command[256];
		if (FAILED(pMenu->GetCommandString(res-CMD_LAST,GCS_VERBA,NULL,command,_countof(command))))
			command[0]=0;
		bool bDelete=false;
		bool bLink=false;
		if (_stricmp(command,"rename")==0)
		{
			// show the Rename dialog box

			s_pDragSource=this; // HACK - prevent the menu from closing
			STRRET str;
			if (SUCCEEDED(pFolder->GetDisplayNameOf(pidl,SHGDN_FOREDITING,&str)))
			{
				wchar_t *name;
				StrRetToStr(&str,pidl,&name);
				g_RenameText=name;
				CoTaskMemFree(name);
			}
			else
				g_RenameText=item.name;
			g_RenamePos=pPt;
			for (std::vector<CMenuContainer*>::iterator it=s_Menus.begin();it!=s_Menus.end();++it)
				(*it)->EnableWindow(FALSE); // disable all menus
			HWND temp=NULL;
			if (!m_pParent)
			{
				// HACK - if we don't create a temp hidden parent, our top menu is brought to foreground
				// this puts it in front of the start button. bad.
				temp=CreateWindow(L"button",NULL,WS_POPUP,0,0,0,0,NULL,NULL,g_Instance,0);
			}

			bool bRenamed=DialogBox(g_Instance,MAKEINTRESOURCE(s_bRTL?IDD_RENAMER:IDD_RENAME),temp?temp:m_hWnd,RenameDlgProc)!=0;

			for (std::vector<CMenuContainer*>::iterator it=s_Menus.begin();it!=s_Menus.end();++it)
				(*it)->EnableWindow(TRUE); // enable all menus
			if (temp)
				::DestroyWindow(temp);
			SetForegroundWindow(m_hWnd);
			SetActiveWindow();
			m_Toolbars[0].SetFocus();
			s_pDragSource=NULL;

			if (bRenamed)
			{
				// perform the rename operation
				PITEMID_CHILD newPidl;
				if (SUCCEEDED(pFolder->SetNameOf(m_hWnd,pidl,g_RenameText,SHGDN_FOREDITING,&newPidl)))
				{
					STRRET str;
					if (SUCCEEDED(pFolder->GetDisplayNameOf(newPidl,SHGDN_INFOLDER|SHGDN_FORPARSING,&str)))
					{
						wchar_t *name;
						StrRetToStr(&str,newPidl,&name);
						CharUpper(name);
						item.name=g_RenameText;
						item.nameHash=CalcFNVHash(name);
						CoTaskMemFree(name);

						std::vector<SortMenuItem> items;
						for (std::vector<MenuItem>::const_iterator it=m_Items.begin();it!=m_Items.end();++it)
							if (it->id==MENU_NO)
							{
								SortMenuItem item={it->name,it->nameHash,it->bFolder};
								items.push_back(item);
							}
						SaveItemOrder(items);
					}
					ILFree(newPidl);
				}
				PostRefreshMessage();
			}
		}
		else if (_stricmp(command,"delete")==0)
			bDelete=true;
		else if (_stricmp(command,"link")==0)
			bLink=true;
		CMINVOKECOMMANDINFOEX info={sizeof(info),CMIC_MASK_UNICODE};
		info.hwnd=m_hWnd;
		info.lpVerb=MAKEINTRESOURCEA(res-CMD_LAST);
		info.lpVerbW=MAKEINTRESOURCEW(res-CMD_LAST);
		info.nShow=SW_SHOWNORMAL;
		if (pPt)
		{
			info.fMask|=CMIC_MASK_PTINVOKE;
			info.ptInvoke=*pPt;
		}
		if (type==ACTIVATE_MENU)
		{
			if (GetKeyState(VK_CONTROL)<0) info.fMask|=CMIC_MASK_CONTROL_DOWN;
			if (GetKeyState(VK_SHIFT)<0) info.fMask|=CMIC_MASK_SHIFT_DOWN;
		}

		HWND temp=NULL;
		if (bDelete)
		{
			// the "delete" verb shows confirmation dialog. we need to prevent the menu from closing
			s_pDragSource=this;
			info.fMask|=CMIC_MASK_NOASYNC;
			for (std::vector<CMenuContainer*>::iterator it=s_Menus.begin();it!=s_Menus.end();++it)
				(*it)->EnableWindow(FALSE);
			if (!m_pParent)
				temp=CreateWindow(L"button",NULL,WS_POPUP,0,0,0,0,NULL,NULL,g_Instance,0);
			info.hwnd=temp?temp:m_hWnd;
		}
		pMenu->InvokeCommand((LPCMINVOKECOMMANDINFO)&info);
		if (bDelete)
		{
			for (std::vector<CMenuContainer*>::iterator it=s_Menus.begin();it!=s_Menus.end();++it)
				(*it)->EnableWindow(TRUE);
			if (temp)
				::DestroyWindow(temp);
			SetForegroundWindow(m_hWnd);
			SetActiveWindow();
			m_Toolbars[0].SetFocus();
			if (bDelete)
				PostRefreshMessage();
			s_pDragSource=NULL;
		}
		else if (bLink)
			PostRefreshMessage(); // just refresh the menu after we create a shortcut
	}
	DestroyMenu(menu);
}

LRESULT CMenuContainer::OnClick( int idCtrl, LPNMHDR pnmh, BOOL& bHandled )
{
	NMMOUSE *pClick=(NMMOUSE*)pnmh;
	int code=pnmh->code;
	if (code==NM_CLICK)
	{
		// check if this was a double click. for some reason the toolbar fails to send NM_DBLCLK even though the docs say it should
		LONG t=GetMessageTime();
		SIZE rect={GetSystemMetrics(SM_CXDOUBLECLK),GetSystemMetrics(SM_CYDOUBLECLK)};
		if (t-m_ClickTime<(LONG)GetDoubleClickTime() && abs(m_ClickPos.x-pClick->pt.x)<=rect.cx/2 && abs(m_ClickPos.y-pClick->pt.y)<=rect.cy/2)
			code=NM_DBLCLK;
		m_ClickTime=t;
		m_ClickPos=pClick->pt;
	}

	POINT pt=pClick->pt;
	ClientToScreen(&pt);

	int index=(int)pClick->dwItemData-ID_OFFSET;
	if (index<0) return 0;
	const MenuItem &item=m_Items[index];

	if (code==NM_DBLCLK || !item.bFolder)
		ActivateItem(index,ACTIVATE_EXECUTE,&pt);
	else if (index!=m_Submenu)
		ActivateItem(index,ACTIVATE_OPEN,NULL);

	return 0;
}

LRESULT CMenuContainer::OnKeyDown( int idCtrl, LPNMHDR pnmh, BOOL& bHandled )
{
	NMKEY *pKey=(NMKEY*)pnmh;
	CWindow toolbar(pnmh->hwndFrom);
	ShowKeyboardCues();

	if (pKey->nVKey!=VK_UP && pKey->nVKey!=VK_DOWN && pKey->nVKey!=VK_LEFT && pKey->nVKey!=VK_RIGHT && pKey->nVKey!=VK_ESCAPE && pKey->nVKey!=VK_RETURN)
		return TRUE;

	int index=(int)toolbar.SendMessage(TB_GETHOTITEM);
	if (index>=0)
	{
		// find the item index from button index
		TBBUTTON button;
		toolbar.SendMessage(TB_GETBUTTON,index,(LPARAM)&button);
		index=(int)button.dwData-ID_OFFSET;
	}
	else
		index=-1;
	int n=(int)m_Items.size();

	if (pKey->nVKey==VK_UP)
	{
		// previous item
		if (index<0) index=0;
		while (m_Items[index].id==MENU_SEPARATOR)
			index=(index+n-1)%n;
		index=(index+n-1)%n;
		while (m_Items[index].id==MENU_SEPARATOR)
			index=(index+n-1)%n;
		ActivateItem(index,ACTIVATE_SELECT,NULL);
	}
	if (pKey->nVKey==VK_DOWN)
	{
		// next item
		while (index>=0 && m_Items[index].id==MENU_SEPARATOR)
			index=(index+1)%n;
		index=(index+1)%n;
		while (m_Items[index].id==MENU_SEPARATOR)
			index=(index+1)%n;
		ActivateItem(index,ACTIVATE_SELECT,NULL);
	}
	if (pKey->nVKey==VK_ESCAPE || (pKey->nVKey==VK_LEFT && !s_bRTL) || (pKey->nVKey==VK_RIGHT && s_bRTL))
	{
		// close top menu
		s_Menus[s_Menus.size()-1]->PostMessage(WM_CLOSE);
		if (s_Menus.size()>=2)
			s_Menus[s_Menus.size()-2]->SetActiveWindow();
	}
	if (pKey->nVKey==VK_RETURN || (pKey->nVKey==VK_RIGHT && !s_bRTL) || (pKey->nVKey==VK_LEFT && s_bRTL))
	{
		// open submenu
		if (index>=0)
		{
			if (m_Items[index].bFolder)
				ActivateItem(index,ACTIVATE_OPEN_KBD,NULL);
			else if (pKey->nVKey==VK_RETURN)
				ActivateItem(index,ACTIVATE_EXECUTE,NULL);
		}
	}
	return TRUE;
}

LRESULT CMenuContainer::OnChar( int idCtrl, LPNMHDR pnmh, BOOL& bHandled )
{
	NMCHAR *pChar=(NMCHAR*)pnmh;

	if (pChar->ch>=0xD800 && pChar->ch<=0xDBFF)
		return TRUE; // don't support supplementary characters

	// find the current menu item
	CWindow toolbar(pnmh->hwndFrom);
	int index=(int)toolbar.SendMessage(TB_GETHOTITEM);
	if (index>=0)
	{
		TBBUTTON button;
		toolbar.SendMessage(TB_GETBUTTON,index,(LPARAM)&button);
		index=(int)button.dwData-ID_OFFSET;
	}
	else
		index=-1;

	// find the next item with that accelerator
	wchar_t buf[2]={pChar->ch,0};
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

	return TRUE;
}

LRESULT CMenuContainer::OnHotItemChange( int idCtrl, LPNMHDR pnmh, BOOL& bHandled )
{
	NMTBHOTITEM *pItem=(NMTBHOTITEM*)pnmh;
	if (m_ContextItem!=-1 && pItem->idNew!=m_ContextItem)
		return TRUE;
	if ((pItem->dwFlags&(HICF_MOUSE|HICF_LEAVING))==HICF_MOUSE)
	{
		if (m_HotPos==GetMessagePos())
			return TRUE; // HACK - ignore the mouse if it hasn't moved since last time. otherwise the mouse can override the keyboard navigation
		if (m_Submenu<0)
			::SetFocus(pnmh->hwndFrom);
		for (std::vector<CWindow>::iterator it=m_Toolbars.begin();it!=m_Toolbars.end();++it)
			if (it->m_hWnd!=pnmh->hwndFrom)
				it->PostMessage(TB_SETHOTITEM,-1);
		if ((m_Submenu>=0 && pItem->idNew-ID_OFFSET!=m_Submenu) || (m_Submenu<0 && m_Items[pItem->idNew-ID_OFFSET].bFolder))
		{
			// initialize the hover timer
			m_HoverItem=pItem->idNew-ID_OFFSET;
			SetTimer(TIMER_HOVER,s_HoverTime);
		}
	}
	if ((pItem->dwFlags&HICF_LEAVING) && m_Submenu>=0)
	{
		// when the mouse moves away, restore the hot item to be the opened submenu
		CWindow column=m_Toolbars[m_Items[m_Submenu].column];
		for (std::vector<CWindow>::iterator it=m_Toolbars.begin();it!=m_Toolbars.end();++it)
			if (it->m_hWnd!=pnmh->hwndFrom && ((int)it->SendMessage(TB_GETHOTITEM))>=0)
				return FALSE;
		column.PostMessage(TB_SETHOTITEM,m_Items[m_Submenu].btnIndex);
		return column.m_hWnd==pnmh->hwndFrom;
	}
	return FALSE;
}

LRESULT CMenuContainer::OnGetInfoTip( int idCtrl, LPNMHDR pnmh, BOOL& bHandled )
{
	if (m_Submenu>=0) return 0; // don't show tips if a submenu is opened

	NMTBGETINFOTIP *pTip=(NMTBGETINFOTIP*)pnmh;
	const MenuItem &item=m_Items[pTip->lParam-ID_OFFSET];
	int idx=-1;
	if (item.id!=MENU_NO && item.id!=MENU_SEPARATOR && item.id!=MENU_EMPTY)
		idx=FindStdMenuItem(item.id);
	if (idx>=0 && g_StdMenu[idx].tipKey)
	{
		// show the tip for the standard item
		wcscpy_s(pTip->pszText,pTip->cchTextMax,FindSetting(g_StdMenu[idx].tipKey,g_StdMenu[idx].tip2));
		return 0;
	}

	// get the tip from the shell
	CComPtr<IShellFolder> pFolder;
	PCUITEMID_CHILD pidl;
	if (FAILED(SHBindToParent(item.pItem1,IID_IShellFolder,(void**)&pFolder,&pidl)))
		return 0;

	CComPtr<IQueryInfo> pQueryInfo;
	if (FAILED(pFolder->GetUIObjectOf(NULL,1,&pidl,IID_IQueryInfo,NULL,(void**)&pQueryInfo)))
		return 0;

	wchar_t *tip;
	if (FAILED(pQueryInfo->GetInfoTip(QITIPF_DEFAULT,&tip)) || !tip)
		return 0;

	wcscpy_s(pTip->pszText,pTip->cchTextMax,tip);
	CoTaskMemFree(tip);
	return 0;
}

// Turn on the keyboard cues from now on. This is done when a keyboard action is detected
void CMenuContainer::ShowKeyboardCues( void )
{
	if (!s_bKeyboardCues)
	{
		s_bKeyboardCues=true;
		for (std::vector<CMenuContainer*>::const_iterator it=s_Menus.begin();it!=s_Menus.end();++it)
		{
			for (std::vector<CWindow>::iterator it2=(*it)->m_Toolbars.begin(),end=(*it)->m_Toolbars.end();it2!=end;++it2)
			{
				it2->SendMessage(TB_SETDRAWTEXTFLAGS,DT_HIDEPREFIX,0);
				it2->InvalidateRect(NULL);
			}
		}
	}
}

void CMenuContainer::SetActiveWindow( void )
{
	if (GetActiveWindow()!=m_hWnd)
		::SetActiveWindow(m_hWnd);
	if (!m_pParent && s_bBehindTaskbar)
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
		// the mouse hovers over an item. open it.
		if (m_HoverItem!=m_Submenu && m_Toolbars[m_Items[m_HoverItem].column].SendMessage(TB_GETHOTITEM)==m_Items[m_HoverItem].btnIndex)
			ActivateItem(m_HoverItem,ACTIVATE_OPEN,NULL);
		KillTimer(TIMER_HOVER);
	}
	return 0;
}

// Handle right-click and the menu keyboard button
LRESULT CMenuContainer::OnContextMenu( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled )
{
	if (s_bNoEditMenu) return 0;
	CWindow toolbar=(HWND)wParam;
	if (toolbar.m_hWnd==m_Pager.m_hWnd)
		toolbar=m_Toolbars[0];
	POINT pt={(short)LOWORD(lParam),(short)HIWORD(lParam)};
	int index;
	if (pt.x!=-1 || pt.y!=-1)
	{
		POINT pt2=pt;
		toolbar.ScreenToClient(&pt2);
		index=(int)toolbar.SendMessage(TB_HITTEST,0,(LPARAM)&pt2);
		if (index<0) return 0;
	}
	else
	{
		index=(int)toolbar.SendMessage(TB_GETHOTITEM);
	}
	TBBUTTON button;
	toolbar.SendMessage(TB_GETBUTTON,index,(LPARAM)&button);
	ActivateItem((int)button.dwData-ID_OFFSET,ACTIVATE_MENU,&pt);

	return 0;
}

// Calculate the size of the toolbar inside the pager
LRESULT CMenuContainer::OnPagerCalcSize( int idCtrl, LPNMHDR pnmh, BOOL& bHandled )
{
	NMPGCALCSIZE *pSize=(NMPGCALCSIZE*)pnmh;
	RECT rc;
	m_Toolbars[0].GetWindowRect(&rc);
	if (pSize->dwFlag&PGF_CALCWIDTH) pSize->iWidth=rc.right-rc.left;
	if (pSize->dwFlag&PGF_CALCHEIGHT) pSize->iHeight=rc.bottom-rc.top;
	return 0;
}

// Adjust the pager scroll speed
LRESULT CMenuContainer::OnPagerScroll( int idCtrl, LPNMHDR pnmh, BOOL& bHandled )
{
	NMPGSCROLL *pScroll=(NMPGSCROLL*)pnmh;
	pScroll->iScroll=HIWORD(m_Toolbars[0].SendMessage(TB_GETBUTTONSIZE));
	return 0;
}

LRESULT CMenuContainer::OnDestroy( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled )
{
	m_bDestroyed=true;
	if (!m_pParent)
	{
		// cleanup when the last menu is closed
		EnableStartTooltip(true);
		if (s_ThemeList) CloseThemeData(s_ThemeList);
		s_ThemeList=NULL;
		if (s_ThemeMenu) CloseThemeData(s_ThemeMenu);
		s_ThemeMenu=NULL;
	}
	return 0;
}

LRESULT CMenuContainer::OnRefresh( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled )
{
	// updates the menu after drag/drop, delete, or rename operation
	m_bRefreshPosted=false;
	InitItems();
	InitToolbars();
	return 0;
}

LRESULT CMenuContainer::OnEraseBkgnd( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled )
{
	// fill the background with COLOR_MENU
	RECT rc;
	GetClientRect(&rc);
	HDC hdc=(HDC)wParam;
	rc.left=m_BmpWidth;
	SetDCBrushColor(hdc,s_MenuColor);
	FillRect(hdc,&rc,(HBRUSH)GetStockObject(DC_BRUSH));

	if (m_Bitmap)
	{
		// draw the caption
		HDC hdc2=CreateCompatibleDC(hdc);
		HGDIOBJ bmp0=SelectObject(hdc2,m_Bitmap);
		BitBlt(hdc,0,0,m_BmpWidth,rc.bottom,hdc2,0,0,SRCCOPY);
		SelectObject(hdc2,bmp0);
		DeleteDC(hdc2);
	}
	return 1;
}

LRESULT CMenuContainer::OnActivate( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled )
{
	if (LOWORD(wParam)!=WA_INACTIVE)
	{
		if (m_Submenu>=0)
			ActivateItem(m_Submenu,ACTIVATE_SELECT,NULL);
		else
			ActivateItem(-1,ACTIVATE_SELECT,NULL);
		return 0;
	}

#ifdef ALLOW_DEACTIVATE
	if (s_pDragSource) return 0;
	// check if another menu window is being activated
	// if not, close all menus
	for (std::vector<CMenuContainer*>::const_iterator it=s_Menus.begin();it!=s_Menus.end();++it)
		if ((*it)->m_hWnd==(HWND)lParam)
			return 0;

	for (std::vector<CMenuContainer*>::reverse_iterator it=s_Menus.rbegin();it!=s_Menus.rend();++it)
		if (!(*it)->m_bDestroyed)
			(*it)->PostMessage(WM_CLOSE);
#endif

	return 0;
}

void CMenuContainer::SaveItemOrder( const std::vector<SortMenuItem> &items )
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

void CMenuContainer::LoadItemOrder( void )
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
		}
	}

	unsigned int hash0=CalcFNVHash(L"");

	// assign each m_Item an index based on its position in items. store in btnIndex
	// unknown items get the index of the blank item, or at the end
	for (std::vector<MenuItem>::iterator it=m_Items.begin();it!=m_Items.end();++it)
	{
		unsigned int hash=it->nameHash;
		it->btnIndex=(int)hashes.size();
		for (size_t i=0;i<hashes.size();i++)
		{
			if (hashes[i]==hash)
			{
				it->btnIndex=(int)i;
				break;
			}
			else if (hashes[i]==hash0)
				it->btnIndex=(int)i;
		}
	}

	// sort by btnIndex, then by bFolder, then by name
	std::sort(m_Items.begin(),m_Items.end());
}

///////////////////////////////////////////////////////////////////////////////

bool CMenuContainer::CloseStartMenu( void )
{
	if (s_Menus.empty()) return false;

	::SetActiveWindow(g_StartButton);

	return true;
}

// Toggles the start menu
HWND CMenuContainer::ToggleStartMenu( HWND startButton, bool bKeyboard )
{
	if (CloseStartMenu())
		return NULL;

	EnableStartTooltip(false);

	// initialize all settings
	StartMenuSettings settings;
	ReadSettings(settings);

	s_bScrollMenus=(settings.ScrollMenus!=0);
	s_bExpandLinks=(settings.ExpandLinks!=0);
	s_MaxRecentDocuments=settings.RecentDocuments;
	s_ShellFormat=RegisterClipboardFormat(CFSTR_SHELLIDLIST);

	bool bRemote=GetSystemMetrics(SM_REMOTESESSION)!=0;
	wchar_t wabPath[_MAX_PATH]=L"%ProgramFiles%\\Windows Mail\\wab.exe";
	DoEnvironmentSubst(wabPath,_countof(wabPath));
	HANDLE hWab=CreateFile(wabPath,0,FILE_SHARE_WRITE,NULL,OPEN_EXISTING,FILE_ATTRIBUTE_NORMAL,NULL);
	bool bPeople=(hWab!=INVALID_HANDLE_VALUE);
	if (bPeople) CloseHandle(hWab);
	s_bRTL=IsLanguageRTL();

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

	bool bNoSetFolders=SHRestricted(REST_NOSETFOLDERS)!=0; // hide control panel, printers, network

	for (int i=0;i<_countof(g_StdOptions);i++)
	{
		switch (g_StdOptions[i].id)
		{
			case MENU_FAVORITES:
				g_StdOptions[i].options=(settings.ShowFavorites && !SHRestricted(REST_NOFAVORITESMENU))?MENU_ENABLED|MENU_EXPANDED:0;
				break;
			case MENU_DOCUMENTS:
				g_StdOptions[i].options=(settings.ShowDocuments && !SHRestricted(REST_NORECENTDOCSMENU))?MENU_ENABLED|MENU_EXPANDED:0;
				break;
			case MENU_LOGOFF:
				{
					DWORD logoff1=SHRestricted(REST_STARTMENULOGOFF);
					DWORD logoff2=SHRestricted(REST_FORCESTARTMENULOGOFF);
					g_StdOptions[i].options=(logoff1!=1 && (logoff1==2 || logoff2 || settings.ShowLogOff))?MENU_ENABLED:0;
				}
				break;
			case MENU_DISCONNECT:
				g_StdOptions[i].options=(bRemote && !SHRestricted(REST_NODISCONNECT))?MENU_ENABLED:0;
				break;
			case MENU_SHUTDOWN:
				g_StdOptions[i].options=(!bRemote && !SHRestricted(REST_NOCLOSE))?MENU_ENABLED:0;
				break;
			case MENU_UNDOCK:
				{
					HW_PROFILE_INFO info;
					GetCurrentHwProfile(&info);
					g_StdOptions[i].options=settings.ShowUndock && ((info.dwDockInfo&(DOCKINFO_DOCKED|DOCKINFO_UNDOCKED))==DOCKINFO_DOCKED && !SHRestricted(REST_NOSMEJECTPC));
				}
				break;
			case MENU_CONTROLPANEL:
				if (bNoSetFolders || SHRestricted(REST_NOCONTROLPANEL))
					g_StdOptions[i].options&=~MENU_ENABLED;
				else
					g_StdOptions[i].options=MENU_ENABLED|(settings.ExpandControlPanel?MENU_EXPANDED:0)|(g_StdOptions[i].options&~MENU_EXPANDED);
				break;
			case MENU_NETWORK:
				if (bNoSetFolders || SHRestricted(REST_NONETWORKCONNECTIONS))
					g_StdOptions[i].options&=~MENU_ENABLED;
				else
					g_StdOptions[i].options=MENU_ENABLED|(settings.ExpandNetwork?MENU_EXPANDED:0)|(g_StdOptions[i].options&~MENU_EXPANDED);
				break;
			case MENU_PRINTERS:
				if (bNoSetFolders)
					g_StdOptions[i].options&=~MENU_ENABLED;
				else
					g_StdOptions[i].options=MENU_ENABLED|(settings.ExpandPrinters?MENU_EXPANDED:0)|(g_StdOptions[i].options&~MENU_EXPANDED);
				break;

			case MENU_SEARCH_PRINTER:
				g_StdOptions[i].options=s_bActiveDirectory==1?MENU_ENABLED:0;
				break;
			case MENU_SEARCH_COMPUTERS:
				g_StdOptions[i].options=(s_bActiveDirectory==1 && !SHRestricted(REST_HASFINDCOMPUTERS))?MENU_ENABLED:0;
				break;
			case MENU_SEARCH_PEOPLE:
				g_StdOptions[i].options=bPeople?MENU_ENABLED:0;
				break;

			case MENU_HELP:
				g_StdOptions[i].options=!SHRestricted(REST_NOSMHELP)?MENU_ENABLED:0;
				break;
			case MENU_RUN:
				g_StdOptions[i].options=!SHRestricted(REST_NORUN)?MENU_ENABLED:0;
				break;
			case MENU_TASKBAR:
				g_StdOptions[i].options=!SHRestricted(REST_NOSETTASKBAR)?MENU_ENABLED:0;
				break;
			case MENU_FEATURES:
				g_StdOptions[i].options=(!bNoSetFolders && !SHRestricted(REST_NOCONTROLPANEL))?MENU_ENABLED:0;
				break;
			case MENU_SEARCH:
				g_StdOptions[i].options=!SHRestricted(REST_NOFIND)?MENU_ENABLED:0;
				break;
			case MENU_USERFILES:
			case MENU_USERDOCUMENTS:
				g_StdOptions[i].options=!SHRestricted(REST_NOSMMYDOCS)?MENU_ENABLED:0;
				break;
			case MENU_USERPICTURES:
				g_StdOptions[i].options=!SHRestricted(REST_NOSMMYPICS)?MENU_ENABLED:0;
				break;
		}
	}	

	s_bNoEditMenu=SHRestricted(REST_NOCHANGESTARMENU)!=0;
	s_bKeyboardCues=bKeyboard;

	// make sure the menu stays on the same monitor as the start button
	RECT startRect;
	::GetWindowRect(startButton,&startRect);
	MONITORINFO info={sizeof(info)};
	GetMonitorInfo(MonitorFromWindow(startButton,MONITOR_DEFAULTTOPRIMARY),&info);
	s_MainRect=info.rcMonitor;

	RECT taskbarRect;
	::GetWindowRect(g_TaskBar,&taskbarRect);
	SystemParametersInfo(SPI_GETMENUSHOWDELAY,NULL,&s_HoverTime,0);

	// create the top menu from the Start Menu folders
	PIDLIST_ABSOLUTE path1;
	PIDLIST_ABSOLUTE path2;
	SHGetKnownFolderIDList(FOLDERID_StartMenu,0,NULL,&path1);
	SHGetKnownFolderIDList(FOLDERID_CommonStartMenu,0,NULL,&path2);

	int options=CONTAINER_PAGER|CONTAINER_NOPROGRAMS|CONTAINER_DRAG|CONTAINER_DROP;
	if (!settings.UseSmallIcons)
		options|=CONTAINER_LARGE;
	if (settings.ConfirmLogOff)
		options|=CONTAINER_CONFIRM_LO;

	DWORD fade;
	SystemParametersInfo(SPI_GETMENUFADE,NULL,&fade,0);
	DWORD animFlags=(fade?AW_BLEND:AW_SLIDE);

	s_bBehindTaskbar=true;
	s_bShowTopEmpty=false;
	RECT margin={0,0,0,0};
	BOOL theme=IsAppThemed();
	if (theme)
	{
		s_MenuBorder=1;
		s_MenuStyle=WS_POPUP|WS_BORDER;
		AdjustWindowRectEx(&margin,s_MenuStyle,FALSE,WS_EX_TOOLWINDOW);
		if (settings.UseTheme) options|=CONTAINER_THEME;
	}
	else
	{
		s_MenuBorder=0;
		s_MenuStyle=WS_POPUP|WS_DLGFRAME;
	}
	RECT rc;
	if (taskbarRect.top>(s_MainRect.top+s_MainRect.bottom)/2)
	{
		// taskbar is at the bottom
		rc.top=rc.bottom=taskbarRect.top+margin.bottom;

		// animate up
		animFlags|=AW_VER_NEGATIVE;
	}
	else if (taskbarRect.bottom<(s_MainRect.top+s_MainRect.bottom)/2)
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

	if (startRect.right<(s_MainRect.left+s_MainRect.right)/2)
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

	CMenuContainer *pStartMenu=new CMenuContainer(NULL,options,g_StdMenu[0].id,path1,path2,L"Software\\IvoSoft\\ClassicStartMenu\\Order");
	pStartMenu->InitItems();
	ILFree(path1);
	ILFree(path2);

	pStartMenu->Create(NULL,&rc,NULL,s_MenuStyle,WS_EX_TOOLWINDOW|WS_EX_TOPMOST|(s_bRTL?WS_EX_LAYOUTRTL:0));
	pStartMenu->m_Toolbars[0].SendMessage(TB_SETHOTITEM,-1);

	BOOL animate;
	if (s_bRTL)
		animate=FALSE; // toolbars don't handle WM_PRINT correctly with RTL, so AnimateWindow doesn't work. disable animations with RTL
	else
		SystemParametersInfo(SPI_GETMENUANIMATION,NULL,&animate,0);

	if (s_bBehindTaskbar)
		::SetWindowPos(startButton,HWND_TOPMOST,0,0,0,0,SWP_NOMOVE|SWP_NOSIZE); // bring the start button on top
	if (animate)
		AnimateWindow(pStartMenu->m_hWnd,MENU_ANIM_SPEED,animFlags);
	else
		pStartMenu->ShowWindow(SW_SHOW);
	pStartMenu->m_Toolbars[0].SetFocus();
	pStartMenu->m_Toolbars[0].SendMessage(TB_SETHOTITEM,-1);
	// position the start button on top
	if (s_bBehindTaskbar)
		::SetWindowPos(startButton,HWND_TOPMOST,0,0,0,0,SWP_NOMOVE|SWP_NOSIZE|SWP_NOACTIVATE);

	return pStartMenu->m_hWnd;
}
