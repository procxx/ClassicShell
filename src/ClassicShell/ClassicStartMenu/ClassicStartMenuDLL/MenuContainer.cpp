// Classic Shell (c) 2009-2010, Ivo Beltchev
// The sources for Classic Shell are distributed under the MIT open source license

#include "stdafx.h"
#include "IconManager.h"
#include "MenuContainer.h"
#include "Accessibility.h"
#include "ClassicStartMenuDLL.h"
#include "Settings.h"
#include "GlobalSettings.h"
#include "ParseSettings.h"
#include "TranslationSettings.h"
#include "CustomMenu.h"
#include "FNVHash.h"
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

#define ALLOW_DEACTIVATE // undefine this to prevent the menu from closing when it is deactivated (useful for debugging)
//#define REPEAT_ITEMS 10 // define this to repeat each menu item (useful to simulate large menus)

#if defined(BUILD_SETUP) && !defined(ALLOW_DEACTIVATE)
#define ALLOW_DEACTIVATE // make sure it is defined in Setup
#endif

#if defined(BUILD_SETUP) && defined(REPEAT_ITEMS)
#undef REPEAT_ITEMS
#endif

const int MAX_MENU_ITEMS=500;
const int MENU_ANIM_SPEED=200;
const int MENU_ANIM_SPEED_SUBMENU=100;
const int MENU_FADE_SPEED=400;

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
	{MENU_UNDOCK,MENU_ENABLED}, // from settings, check policy
	{MENU_CONTROLPANEL,MENU_ENABLED|MENU_EXPANDED}, // MENU_EXPANDED from settings, check policy
	{MENU_NETWORK,MENU_ENABLED}, // MENU_EXPANDED from settings, check policy
	{MENU_PRINTERS,MENU_ENABLED}, // MENU_EXPANDED from settings, check policy
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
bool CMenuContainer::s_bBehindTaskbar=true;
bool CMenuContainer::s_bShowTopEmpty=false;
bool CMenuContainer::s_bNoEditMenu=false;
bool CMenuContainer::s_bExpandLinks=false;
char CMenuContainer::s_bActiveDirectory=-1;
CMenuContainer *CMenuContainer::s_pDragSource=NULL;
bool CMenuContainer::s_bRightDrag;
std::vector<CMenuContainer*> CMenuContainer::s_Menus;
std::map<unsigned int,int> CMenuContainer::s_PagerScrolls;
CComPtr<IShellFolder> CMenuContainer::s_pDesktop;
HWND CMenuContainer::s_LastFGWindow;
RECT CMenuContainer::s_MainRect;
DWORD CMenuContainer::s_TaskbarState;
DWORD CMenuContainer::s_HoverTime;
DWORD CMenuContainer::s_SubmenuStyle;
CLIPFORMAT CMenuContainer::s_ShellFormat;
MenuSkin CMenuContainer::s_Skin;
std::vector<CMenuFader*> CMenuFader::s_Faders;

// Subclass proc for the toolbars
LRESULT CALLBACK CMenuContainer::ToolbarSubclassProc( HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData )
{
	CMenuContainer *pOwner=(CMenuContainer*)uIdSubclass;

	if (uMsg==WM_MOUSEACTIVATE)
		return MA_NOACTIVATE; // prevent activation with the mouse
	if (uMsg==WM_GETOBJECT && (DWORD)lParam==(DWORD)OBJID_CLIENT && pOwner->m_pAccessible)
		return LresultFromObject(IID_IAccessible,wParam,pOwner->m_pAccessible);

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
		LRESULT res=0;
		pOwner->AddRef(); // the owner may be deleted during DefSubclassProc
		if (idx>=0 || GetCapture()==hWnd)
			res=DefSubclassProc(hWnd,uMsg,wParam,lParam);
		pOwner->m_HotPos=GetMessagePos();
		pOwner->Release();
		return res;
	}
	if (uMsg==WM_PAINT && !pOwner->m_pParent && s_Skin.Main_opacity>=MenuSkin::OPACITY_ALPHA)
	{
		PAINTSTRUCT ps;
		HDC hdc=::BeginPaint(hWnd,&ps);
		if(hdc)
		{
			BP_PAINTPARAMS paintParams = {0};
			paintParams.cbSize = sizeof(paintParams);
			paintParams.dwFlags=BPPF_ERASE;

			HDC hdcPaint=NULL;
			HPAINTBUFFER hBufferedPaint=BeginBufferedPaint(hdc,&ps.rcPaint,BPBF_TOPDOWNDIB,&paintParams,&hdcPaint);
			if (hdcPaint)
			{
				DefSubclassProc(hWnd,WM_PRINTCLIENT,(WPARAM)hdcPaint,PRF_CLIENT);
				BufferedPaintSetAlpha(hBufferedPaint,&ps.rcPaint,255);
				EndBufferedPaint(hBufferedPaint,TRUE);
			}
		}
		::EndPaint(hWnd,&ps);
		return 0;
	}
	return DefSubclassProc(hWnd,uMsg,wParam,lParam);
}

CMenuContainer::CMenuContainer( CMenuContainer *pParent, int index, int options, const StdMenuItem *pStdItem, PIDLIST_ABSOLUTE path1, PIDLIST_ABSOLUTE path2, const CString &regName )
{
	m_RefCount=1;
	m_pParent=pParent;
	m_ParentIndex=index;
	m_Options=options;
	m_pStdItem=pStdItem;
	m_RegName=regName;
	m_Bitmap=NULL;
	m_Region=NULL;
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
	{
		if (m_pParent->m_Submenu==m_ParentIndex)
			m_pParent->m_Submenu=-1;
	}
	if (m_Bitmap) DeleteObject(m_Bitmap);
	if (m_Region) DeleteObject(m_Region);

	// must be here and not in OnDestroy because during drag/drop a menu can close while still processing messages
	s_Menus.erase(std::find(s_Menus.begin(),s_Menus.end(),this));
}

// Initialize the m_Items list
void CMenuContainer::InitItems( void )
{
	m_Items.clear();
	m_bRefreshPosted=false;
	m_Submenu=-1;
	m_HotPos=GetMessagePos();
	m_ContextItem=-1;

	if ((m_Options&CONTAINER_DOCUMENTS) && s_MaxRecentDocuments>0) // create the recent documents list
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
				item.bPrograms=(m_Options&CONTAINER_PROGRAMS)!=0;

				m_Items.push_back(item);
#ifdef REPEAT_ITEMS
				for (int i=0;i<REPEAT_ITEMS;i++)
				{
					item.pItem1=ILCloneFull(item.pItem1);
					m_Items.push_back(item);
				}
#endif
			}
			ILFree(pidl);
		}
	}

	// add second folder
	if (m_Path2 && !(m_Options&CONTAINER_DOCUMENTS))
	{
		CComPtr<IShellFolder> pFolder;
		s_pDesktop->BindToObject(m_Path2,NULL,IID_IShellFolder,(void**)&pFolder);

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
					item.bPrograms=(m_Options&CONTAINER_PROGRAMS)!=0;

					m_Items.push_back(item);
#ifdef REPEAT_ITEMS
					for (int i=0;i<REPEAT_ITEMS;i++)
					{
						item.pItem1=ILCloneFull(item.pItem1);
						m_Items.push_back(item);
					}
#endif
					CoTaskMemFree(name);
				}
			}
			ILFree(pidl);
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

	if (m_Items.empty() && m_Path1 && m_pDropFolder)
	{
		// add (Empty) item to the empty submenus
		MenuItem item={MENU_EMPTY};
		item.icon=I_IMAGENONE;
		item.name=FindTranslation("Menu.Empty",L"(Empty)");
		m_Items.push_back(item);
	}

	// add standard items
	if (m_pStdItem && m_pStdItem->id!=MENU_NO)
	{
		if (!m_Items.empty())
		{
			MenuItem item={MENU_SEPARATOR};
			if (m_Options&CONTAINER_ADDTOP)
				m_Items.insert(m_Items.begin(),item);
			else
				m_Items.push_back(item);
		}
		size_t menuIdx=(m_Options&CONTAINER_ADDTOP)?0:m_Items.size();
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

			MenuItem item={pItem->id,pItem};

			ATLASSERT(pItem->folder1 || !pItem->folder2);
			if (pItem->folder1)
			{
				SHGetKnownFolderIDList(*pItem->folder1,0,NULL,&item.pItem1);
				if (pItem->folder2)
					SHGetKnownFolderIDList(*pItem->folder2,0,NULL,&item.pItem2);
				item.bFolder=((stdOptions&MENU_EXPANDED) || pItem->submenu);
			}
			else if (pItem->link)
			{
				SFGAOF flags=SFGAO_FOLDER|SFGAO_STREAM|SFGAO_LINK;
				wchar_t buf[1024];
				wcscpy_s(buf,item.pStdItem->link);
				DoEnvironmentSubst(buf,_countof(buf));
				if (SUCCEEDED(s_pDesktop->ParseDisplayName(NULL,NULL,buf,NULL,(PIDLIST_RELATIVE*)&item.pItem1,&flags)))
				{
					item.bLink=(flags&SFGAO_LINK)!=0;
					item.bFolder=((flags&SFGAO_FOLDER) && (!(flags&(SFGAO_STREAM|SFGAO_LINK)) || (s_bExpandLinks && item.bLink)));
				}
			}
			if (pItem->submenu)
				item.bFolder=true;

			// get icon
			if (pItem->iconPath)
				item.icon=g_IconManager.GetCustomIcon(pItem->iconPath,(m_Options&CONTAINER_LARGE)!=0);
			else if (pItem->icon)
				item.icon=g_IconManager.GetStdIcon(pItem->icon,(m_Options&CONTAINER_LARGE)!=0);
			else if (item.pItem1)
			{
				// for some reason SHGetFileInfo(SHGFI_PIDL|SHGFI_ICONLOCATION) doesn't work here. go to the IShellFolder directly
				CComPtr<IShellFolder> pFolder2;
				PCUITEMID_CHILD pidl;
				if (SUCCEEDED(SHBindToParent(item.pItem1,IID_IShellFolder,(void**)&pFolder2,&pidl)))
					item.icon=g_IconManager.GetIcon(pFolder2,pidl,(m_Options&CONTAINER_LARGE)!=0);
			}

			// get name
			if (pItem->name)
			{
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
					item.name.Format(pItem->key?FindTranslation(pItem->key,pItem->name):pItem->name,user);
				}
				else
					item.name=pItem->key?FindTranslation(pItem->key,pItem->name):pItem->name;
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

static void MarginsBlit( HDC hSrc, HDC hDst, const RECT &rSrc, const RECT &rDst, const RECT &rMargins )
{
	int x0a=rDst.left;
	int x1a=rDst.left+rMargins.left;
	int x2a=rDst.right-rMargins.right;
	int x3a=rDst.right;
	int x0b=rSrc.left;
	int x1b=rSrc.left+rMargins.left;
	int x2b=rSrc.right-rMargins.right;
	int x3b=rSrc.right;

	int y0a=rDst.top;
	int y1a=rDst.top+rMargins.top;
	int y2a=rDst.bottom-rMargins.bottom;
	int y3a=rDst.bottom;
	int y0b=rSrc.top;
	int y1b=rSrc.top+rMargins.top;
	int y2b=rSrc.bottom-rMargins.bottom;
	int y3b=rSrc.bottom;

	SetStretchBltMode(hDst,COLORONCOLOR);
	if (x0a<x1a && y0a<y1a && x0b<x1b && y0b<y1b) StretchBlt(hDst,x0a,y0a,x1a-x0a,y1a-y0a,hSrc,x0b,y0b,x1b-x0b,y1b-y0b,SRCCOPY);
	if (x1a<x2a && y0a<y1a && x1b<x2b && y0b<y1b) StretchBlt(hDst,x1a,y0a,x2a-x1a,y1a-y0a,hSrc,x1b,y0b,x2b-x1b,y1b-y0b,SRCCOPY);
	if (x2a<x3a && y0a<y1a && x2b<x3b && y0b<y1b) StretchBlt(hDst,x2a,y0a,x3a-x2a,y1a-y0a,hSrc,x2b,y0b,x3b-x2b,y1b-y0b,SRCCOPY);
	
	if (x0a<x1a && y1a<y2a && x0b<x1b && y1b<y2b) StretchBlt(hDst,x0a,y1a,x1a-x0a,y2a-y1a,hSrc,x0b,y1b,x1b-x0b,y2b-y1b,SRCCOPY);
	if (x1a<x2a && y1a<y2a && x1b<x2b && y1b<y2b) StretchBlt(hDst,x1a,y1a,x2a-x1a,y2a-y1a,hSrc,x1b,y1b,x2b-x1b,y2b-y1b,SRCCOPY);
	if (x2a<x3a && y1a<y2a && x2b<x3b && y1b<y2b) StretchBlt(hDst,x2a,y1a,x3a-x2a,y2a-y1a,hSrc,x2b,y1b,x3b-x2b,y2b-y1b,SRCCOPY);
	
	if (x0a<x1a && y2a<y3a && x0b<x1b && y2b<y3b) StretchBlt(hDst,x0a,y2a,x1a-x0a,y3a-y2a,hSrc,x0b,y2b,x1b-x0b,y3b-y2b,SRCCOPY);
	if (x1a<x2a && y2a<y3a && x1b<x2b && y2b<y3b) StretchBlt(hDst,x1a,y2a,x2a-x1a,y3a-y2a,hSrc,x1b,y2b,x2b-x1b,y3b-y2b,SRCCOPY);
	if (x2a<x3a && y2a<y3a && x2b<x3b && y2b<y3b) StretchBlt(hDst,x2a,y2a,x3a-x2a,y3a-y2a,hSrc,x2b,y2b,x3b-x2b,y3b-y2b,SRCCOPY);
}

static void MarginsBlitAlpha( HDC hSrc, HDC hDst, const RECT &rSrc, const RECT &rDst, const RECT &rMargins )
{
	int x0a=rDst.left;
	int x1a=rDst.left+rMargins.left;
	int x2a=rDst.right-rMargins.right;
	int x3a=rDst.right;
	int x0b=rSrc.left;
	int x1b=rSrc.left+rMargins.left;
	int x2b=rSrc.right-rMargins.right;
	int x3b=rSrc.right;

	int y0a=rDst.top;
	int y1a=rDst.top+rMargins.top;
	int y2a=rDst.bottom-rMargins.bottom;
	int y3a=rDst.bottom;
	int y0b=rSrc.top;
	int y1b=rSrc.top+rMargins.top;
	int y2b=rSrc.bottom-rMargins.bottom;
	int y3b=rSrc.bottom;

	SetStretchBltMode(hDst,COLORONCOLOR);
	BLENDFUNCTION func={AC_SRC_OVER,0,255,AC_SRC_ALPHA};
	if (x0a<x1a && y0a<y1a && x0b<x1b && y0b<y1b) AlphaBlend(hDst,x0a,y0a,x1a-x0a,y1a-y0a,hSrc,x0b,y0b,x1b-x0b,y1b-y0b,func);
	if (x1a<x2a && y0a<y1a && x1b<x2b && y0b<y1b) AlphaBlend(hDst,x1a,y0a,x2a-x1a,y1a-y0a,hSrc,x1b,y0b,x2b-x1b,y1b-y0b,func);
	if (x2a<x3a && y0a<y1a && x2b<x3b && y0b<y1b) AlphaBlend(hDst,x2a,y0a,x3a-x2a,y1a-y0a,hSrc,x2b,y0b,x3b-x2b,y1b-y0b,func);

	if (x0a<x1a && y1a<y2a && x0b<x1b && y1b<y2b) AlphaBlend(hDst,x0a,y1a,x1a-x0a,y2a-y1a,hSrc,x0b,y1b,x1b-x0b,y2b-y1b,func);
	if (x1a<x2a && y1a<y2a && x1b<x2b && y1b<y2b) AlphaBlend(hDst,x1a,y1a,x2a-x1a,y2a-y1a,hSrc,x1b,y1b,x2b-x1b,y2b-y1b,func);
	if (x2a<x3a && y1a<y2a && x2b<x3b && y1b<y2b) AlphaBlend(hDst,x2a,y1a,x3a-x2a,y2a-y1a,hSrc,x2b,y1b,x3b-x2b,y2b-y1b,func);

	if (x0a<x1a && y2a<y3a && x0b<x1b && y2b<y3b) AlphaBlend(hDst,x0a,y2a,x1a-x0a,y3a-y2a,hSrc,x0b,y2b,x1b-x0b,y3b-y2b,func);
	if (x1a<x2a && y2a<y3a && x1b<x2b && y2b<y3b) AlphaBlend(hDst,x1a,y2a,x2a-x1a,y3a-y2a,hSrc,x1b,y2b,x2b-x1b,y3b-y2b,func);
	if (x2a<x3a && y2a<y3a && x2b<x3b && y2b<y3b) AlphaBlend(hDst,x2a,y2a,x3a-x2a,y3a-y2a,hSrc,x2b,y2b,x3b-x2b,y3b-y2b,func);
}

// Creates the bitmap to show on the main menu
void CMenuContainer::CreateBackground( int width, int height )
{
	// get the text from the ini file or from the registry
	CRegKey regTitle;
	wchar_t title[256]=L"Windows";
	const wchar_t *setting=FindSetting("MenuCaption");
	if (setting)
		wcscpy_s(title,setting);
	else
	{
		ULONG size=_countof(title);
		if (regTitle.Open(HKEY_LOCAL_MACHINE,L"Software\\Microsoft\\Windows NT\\CurrentVersion",KEY_READ)==ERROR_SUCCESS)
			regTitle.QueryStringValue(L"ProductName",title,&size);
	}

	HBITMAP bmpSkin=s_Skin.Main_bitmap;
	bool b32=s_Skin.Main_bitmap32;
	const int *slicesX=s_Skin.Main_bitmap_slices_X;
	const int *slicesY=s_Skin.Main_bitmap_slices_Y;
	bool bCaption=(slicesX[1]>0);

	HDC hdcTemp=CreateCompatibleDC(NULL);

	HFONT font0=NULL;
	if (bCaption)
		font0=(HFONT)SelectObject(hdcTemp,s_Skin.Caption_font);

	RECT rc={0,0,0,0};
	DTTOPTS opts={sizeof(opts),DTT_COMPOSITED|DTT_CALCRECT};
	HTHEME theme=NULL;
	if (bCaption)
	{
		theme=OpenThemeData(m_hWnd,L"window"); // create a dummy theme to be used by DrawThemeTextEx
		if (theme)
			DrawThemeTextEx(theme,hdcTemp,0,0,title,-1,DT_NOPREFIX|DT_SINGLELINE|DT_CALCRECT,&rc,&opts);
		else
			DrawText(hdcTemp,title,-1,&rc,DT_NOPREFIX|DT_SINGLELINE|DT_CALCRECT);
	}
	int textWidth=rc.right+s_Skin.Caption_padding.top+s_Skin.Caption_padding.bottom;
	int textHeight=rc.bottom+s_Skin.Caption_padding.left+s_Skin.Caption_padding.right;
	int total=slicesX[0]+slicesX[2];
	if (textHeight<total) textHeight=total;

	int totalWidth=textHeight+width+s_Skin.Main_padding.left+s_Skin.Main_padding.right;
	total+=textHeight+slicesX[3]+slicesX[5];
	if (totalWidth<total) totalWidth=total;

	int totalHeight=height+s_Skin.Main_padding.top+s_Skin.Main_padding.bottom;
	total=slicesY[0]+slicesY[2];
	if (totalHeight<total) totalHeight=total;
	if (textWidth>totalHeight) textWidth=totalHeight;

	BITMAPINFO dib={sizeof(dib)};
	dib.bmiHeader.biWidth=totalWidth;
	dib.bmiHeader.biHeight=-totalHeight;
	dib.bmiHeader.biPlanes=1;
	dib.bmiHeader.biBitCount=32;
	dib.bmiHeader.biCompression=BI_RGB;

	HDC hdc=CreateCompatibleDC(NULL);
	unsigned int *bits;
	m_Bitmap=CreateDIBSection(hdc,&dib,DIB_RGB_COLORS,(void**)&bits,NULL,0);
	HBITMAP bmp0=(HBITMAP)SelectObject(hdc,m_Bitmap);

	if (s_Skin.Main_opacity==MenuSkin::OPACITY_SOLID)
	{
		RECT rc={0,0,totalWidth,totalHeight};
		SetDCBrushColor(hdc,s_Skin.Main_background);
		FillRect(hdc,&rc,(HBRUSH)GetStockObject(DC_BRUSH));
	}

	if (bmpSkin)
	{
		// draw the skinned background
		HBITMAP bmp02=(HBITMAP)SelectObject(hdcTemp,bmpSkin);

		RECT rSrc={0,0,slicesX[0]+slicesX[1]+slicesX[2],slicesY[0]+slicesY[1]+slicesY[2]};
		RECT rDst={0,0,textHeight,totalHeight};
		RECT rMargins={slicesX[0],slicesY[0],slicesX[2],slicesY[2]};
		if (s_Skin.Main_opacity==MenuSkin::OPACITY_SOLID && b32)
			MarginsBlitAlpha(hdcTemp,hdc,rSrc,rDst,rMargins);
		else
			MarginsBlit(hdcTemp,hdc,rSrc,rDst,rMargins);

		rSrc.left=rSrc.right;
		rSrc.right+=slicesX[3]+slicesX[4]+slicesX[5];
		rDst.left=rDst.right;
		rDst.right=totalWidth;
		rMargins.left=slicesX[3];
		rMargins.right=slicesX[5];
		if (s_Skin.Main_opacity==MenuSkin::OPACITY_SOLID && b32)
			MarginsBlitAlpha(hdcTemp,hdc,rSrc,rDst,rMargins);
		else
			MarginsBlit(hdcTemp,hdc,rSrc,rDst,rMargins);

		SelectObject(hdcTemp,bmp02);
		SelectObject(hdc,bmp0); // deselect m_Bitmap so all the GDI operations get flushed

		if (s_bRTL)
		{
			// mirror the background image for RTL windows
			for (int y=0;y<totalHeight;y++)
			{
				int yw=y*totalWidth;
				std::reverse(bits+yw,bits+yw+totalWidth);
			}
		}

		// calculate the window region
		m_Region=NULL;
		if (s_Skin.Main_opacity==MenuSkin::OPACITY_REGION || s_Skin.Main_opacity==MenuSkin::OPACITY_GLASS)
		{
			for (int y=0;y<totalHeight;y++)
			{
				int minx=-1, maxx=-1;
				int yw=y*totalWidth;
				for (int x=0;x<totalWidth;x++)
				{
					if (bits[yw+x]&0xFF000000)
					{
						if (minx==-1) minx=x; // first non-transparent pixel
						if (maxx<x) maxx=x; // last non-transparent pixel
					}
				}
				if (minx>=0)
				{
					maxx++;
					if (s_bRTL && s_Skin.Main_opacity==MenuSkin::OPACITY_REGION)
						minx=totalWidth-minx, maxx=totalWidth-maxx; // in "region" mode mirror the region (again)
					HRGN r=CreateRectRgn(minx,y,maxx,y+1);
					if (!m_Region)
						m_Region=r;
					else
					{
						CombineRgn(m_Region,m_Region,r,RGN_OR);
						DeleteObject(r);
					}
				}
			}
		}

		SelectObject(hdc,m_Bitmap);
	}
	else
	{
		RECT rc={0,0,totalWidth,totalHeight};
		SetDCBrushColor(hdc,s_Skin.Main_background);
		FillRect(hdc,&rc,(HBRUSH)GetStockObject(DC_BRUSH));
	}

	if (bCaption)
	{
		// draw the title
		BITMAPINFO dib={sizeof(dib)};
		dib.bmiHeader.biWidth=textWidth;
		dib.bmiHeader.biHeight=-textHeight;
		dib.bmiHeader.biPlanes=1;
		dib.bmiHeader.biBitCount=32;
		dib.bmiHeader.biCompression=BI_RGB;

		HDC hdc=CreateCompatibleDC(NULL);
		unsigned int *bits2;
		HBITMAP bmpText=CreateDIBSection(hdcTemp,&dib,DIB_RGB_COLORS,(void**)&bits2,NULL,0);
		HBITMAP bmp02=(HBITMAP)SelectObject(hdcTemp,bmpText);
		{
			RECT rc={0,0,textWidth,textHeight};
			FillRect(hdcTemp,&rc,(HBRUSH)GetStockObject(BLACK_BRUSH));
		}

		RECT rc={s_Skin.Caption_padding.bottom,s_bRTL?s_Skin.Caption_padding.right:s_Skin.Caption_padding.left,textWidth-s_Skin.Caption_padding.top,textHeight-(s_bRTL?s_Skin.Caption_padding.left:s_Skin.Caption_padding.right)};
		if (theme && s_Skin.Caption_glow_size>0)
		{
			// draw the glow
			opts.dwFlags=DTT_COMPOSITED|DTT_TEXTCOLOR|DTT_GLOWSIZE;
			opts.crText=0xFFFFFF;
			opts.iGlowSize=s_Skin.Caption_glow_size;
			DrawThemeTextEx(theme,hdcTemp,0,0,title,-1,DT_VCENTER|DT_NOPREFIX|DT_SINGLELINE,&rc,&opts);
			SelectObject(hdcTemp,bmp02); // deselect bmpText so all the GDI operations get flushed

			// change the glow color
			int gr=(s_Skin.Caption_glow_color>>16)&255;
			int gg=(s_Skin.Caption_glow_color>>8)&255;
			int gb=(s_Skin.Caption_glow_color)&255;
			for (int y=0;y<textHeight;y++)
				for (int x=0;x<textWidth;x++)
				{
					unsigned int &pixel=bits2[y*textWidth+x];
					int a1=(pixel>>24);
					int r1=(pixel>>16)&255;
					int g1=(pixel>>8)&255;
					int b1=(pixel)&255;
					r1=(r1*gr)/255;
					g1=(g1*gg)/255;
					b1=(b1*gb)/255;
					pixel=(a1<<24)|(r1<<16)|(g1<<8)|b1;
				}

			SelectObject(hdcTemp,bmpText);
		}

		// draw the text
		int offset=0;
		if (s_bRTL)
			offset=totalWidth-textHeight;

		if (theme)
		{
			opts.dwFlags=DTT_COMPOSITED|DTT_TEXTCOLOR;
			opts.crText=s_Skin.Caption_text_color;
			DrawThemeTextEx(theme,hdcTemp,0,0,title,-1,DT_VCENTER|DT_NOPREFIX|DT_SINGLELINE,&rc,&opts);
			SelectObject(hdcTemp,bmp02);

			// rotate and copy the text onto the final bitmap. Combine the alpha channels
			for (int y=0;y<textHeight;y++)
				for (int x=0;x<textWidth;x++)
				{
					unsigned int src=bits2[y*textWidth+x];
					int a1=(src>>24);
					int r1=(src>>16)&255;
					int g1=(src>>8)&255;
					int b1=(src)&255;

					unsigned int &dst=bits[(totalHeight-1-x)*totalWidth+y+offset];

					int a2=(dst>>24);
					int r2=(dst>>16)&255;
					int g2=(dst>>8)&255;
					int b2=(dst)&255;

					r2=(r2*(255-a1))/255+r1;
					g2=(g2*(255-a1))/255+g1;
					b2=(b2*(255-a1))/255+b1;
					a2=a1+a2-(a1*a2)/255;

					dst=(a2<<24)|(r2<<16)|(g2<<8)|b2;
				}
		}
		else
		{
			// draw white text on black background
			FillRect(hdcTemp,&rc,(HBRUSH)GetStockObject(BLACK_BRUSH));
			SetTextColor(hdcTemp,0xFFFFFF);
			SetBkMode(hdcTemp,TRANSPARENT);
			DrawText(hdcTemp,title,-1,&rc,DT_VCENTER|DT_NOPREFIX|DT_SINGLELINE);
			SelectObject(hdcTemp,bmp02);

			// rotate and copy the text onto the final bitmap
			// change the text color
			int tr=(s_Skin.Caption_text_color>>16)&255;
			int tg=(s_Skin.Caption_text_color>>8)&255;
			int tb=(s_Skin.Caption_text_color)&255;
			for (int y=0;y<textHeight;y++)
				for (int x=0;x<textWidth;x++)
				{
					unsigned int src=bits2[y*textWidth+x];
					int a1=(src)&255;

					unsigned int &dst=bits[(totalHeight-1-x)*totalWidth+y+offset];

					int a2=(dst>>24);
					int r2=(dst>>16)&255;
					int g2=(dst>>8)&255;
					int b2=(dst)&255;

					r2=(r2*(255-a1)+tr*a1)/255;
					g2=(g2*(255-a1)+tg*a1)/255;
					b2=(b2*(255-a1)+tb*a1)/255;
					a2=a1+a2-(a1*a2)/255;

					dst=(a2<<24)|(r2<<16)|(g2<<8)|b2;
				}
		}

		DeleteObject(bmpText);
	}

	SelectObject(hdcTemp,font0);
	DeleteDC(hdcTemp);

	SelectObject(hdc,bmp0);
	DeleteDC(hdc);

	if (theme) CloseThemeData(theme);

	m_rContent.left=s_Skin.Main_padding.left+textHeight;
	m_rContent.right=totalWidth-s_Skin.Main_padding.right;
	m_rContent.top=s_Skin.Main_padding.top;
	m_rContent.bottom=totalHeight-s_Skin.Main_padding.bottom;
}

CWindow CMenuContainer::CreateToolbar( int id )
{
	CWindow toolbar=CreateWindowEx(WS_EX_TOOLWINDOW,TOOLBARCLASSNAME,L"",WS_CHILD|WS_VISIBLE|TBSTYLE_TOOLTIPS|TBSTYLE_FLAT|TBSTYLE_LIST|TBSTYLE_TRANSPARENT|TBSTYLE_REGISTERDROP|CCS_NODIVIDER|CCS_NOPARENTALIGN|CCS_NORESIZE|CCS_VERT,0,0,10,10,m_hWnd,(HMENU)id,g_Instance,NULL);
	SetWindowTheme(toolbar.m_hWnd,L"",L""); // disable cross-fade on Vista
	toolbar.SendMessage(TB_BUTTONSTRUCTSIZE,sizeof(TBBUTTON),0);

	HIMAGELIST images=(m_Options&CONTAINER_LARGE)?g_IconManager.m_LargeIcons:g_IconManager.m_SmallIcons;
	toolbar.SendMessage(TB_SETIMAGELIST,0,(LPARAM)images);

	toolbar.SendMessage(TB_SETINSERTMARKCOLOR,0,m_pParent?s_Skin.Submenu_text_color[0]:s_Skin.Main_text_color[0]);
	HFONT font=m_pParent?s_Skin.Submenu_font:s_Skin.Main_font;
	if (font) toolbar.SetFont(font);

	CWindow tooltip=(HWND)toolbar.SendMessage(TB_GETTOOLTIPS);
	if (tooltip.m_hWnd)
	{
		const wchar_t *str=FindSetting("InfotipDelay");
		if (str)
		{
			wchar_t token[256];
			str=GetToken(str,token,_countof(token),L", \t");
			int time=_wtol(token);
			if (time<=0) time=-1;
			tooltip.SendMessage(TTM_SETDELAYTIME,TTDT_INITIAL,time);
			str=GetToken(str,token,_countof(token),L", \t");
			time=_wtol(token);
			if (time<=0) time=-1;
			tooltip.SendMessage(TTM_SETDELAYTIME,TTDT_AUTOPOP,time);
			str=GetToken(str,token,_countof(token),L", \t");
			time=_wtol(token);
			if (time<=0) time=-1;
			tooltip.SendMessage(TTM_SETDELAYTIME,TTDT_RESHOW,time);
		}
/*
see the comment in OnGetInfoTip
		tooltip.SetWindowLong(GWL_STYLE,tooltip.GetWindowLong(GWL_STYLE)|TTS_ALWAYSTIP);
*/
	}

	SetWindowSubclass(toolbar.m_hWnd,ToolbarSubclassProc,(UINT_PTR)this,0);

	return toolbar;
}

// Create one or more toolbars and the pager. Reuses the existing controls if possible
void CMenuContainer::InitToolbars( void )
{
	// calculate maximum height
	int maxHeight;
	{
		maxHeight=(s_MainRect.bottom-s_MainRect.top);
		// adjust for padding
		RECT rc={0,0,0,0};
		AdjustWindowRect(&rc,GetWindowLong(GWL_STYLE),FALSE);
		if (m_pParent)
		{
			maxHeight-=rc.bottom-rc.top;
			maxHeight-=s_Skin.Submenu_padding.top+s_Skin.Submenu_padding.bottom;
		}
		else
		{
			RECT rc2;
			GetWindowRect(&rc2);
			if (m_Options&CONTAINER_TOP)
				maxHeight=s_MainRect.bottom-rc2.top;
			else
				maxHeight=rc2.bottom-s_MainRect.top;
			maxHeight-=rc.bottom-rc.top;
			maxHeight-=s_Skin.Main_padding.top+s_Skin.Main_padding.bottom;
		}
	}
#ifdef _DEBUG
//	maxHeight/=4; // uncomment to test for smaller screen
#endif

	// add buttons
	CWindow toolbar;
	std::vector<CWindow> oldToolbars;
	oldToolbars.swap(m_Toolbars);
	int column=0;
	int sepHeight=0;
	if (m_pParent)
	{
		if (s_Skin.Submenu_separator) sepHeight=s_Skin.Submenu_separatorHeight;
	}
	else
	{
		if (s_Skin.Main_separator) sepHeight=s_Skin.Main_separatorHeight;
	}
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
		{
			button.fsStyle=BTNS_SEP;
			button.iBitmap=sepHeight;
		}
		else if (m_Items[i].id==MENU_NO)
			button.fsStyle|=BTNS_NOPREFIX;
		if (!s_bShowTopEmpty && !m_pParent && (m_Items[i].id==MENU_EMPTY || (i==1 && m_Items[0].id==MENU_EMPTY)))
			button.fsState|=TBSTATE_HIDDEN; // this is the first (Empty) item in the top menu. hide it for now
		int n=(int)toolbar.SendMessage(TB_BUTTONCOUNT);
		m_Items[i].column=column;
		m_Items[i].btnIndex=n;
		toolbar.SendMessage(TB_INSERTBUTTON,n,(LPARAM)&button);

		if (m_Options&CONTAINER_MULTICOLUMN)
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

//	CComPtr<IAccPropServices> pAccPropServices;
//	CoCreateInstance(CLSID_AccPropServices,NULL,CLSCTX_SERVER,IID_IAccPropServices,(void**)&pAccPropServices);
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
/*		if (pAccPropServices)
		{
			pAccPropServices->SetHwndProp(it->m_hWnd,OBJID_CLIENT,CHILDID_SELF,PROPID_ACC_ROLE,CComVariant(ROLE_SYSTEM_MENUPOPUP));
			for (int i=0;i<n;i++)
				pAccPropServices->SetHwndProp(it->m_hWnd,OBJID_CLIENT,i+1,PROPID_ACC_ROLE,CComVariant(ROLE_SYSTEM_MENUITEM));
		}*/
	}

	maxw+=16; // add 16 pixels for the submenu arrow

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
	{
		int w=(int)m_Toolbars.size()*maxw;
		int h=(maxh<maxHeight?maxh:maxHeight);
		if (!m_pParent)
		{
			if (s_Skin.Main_bitmap)
			{
				CreateBackground((int)m_Toolbars.size()*maxw,maxh<maxHeight?maxh:maxHeight);
				BITMAP info;
				GetObject(m_Bitmap,sizeof(info),&info);
				totalWidth=info.bmWidth;
				totalHeight=info.bmHeight;
			}
			else
			{
				m_rContent.left=s_Skin.Main_padding.left;
				m_rContent.top=s_Skin.Main_padding.top;
				m_rContent.right=s_Skin.Main_padding.left+w;
				m_rContent.bottom=s_Skin.Main_padding.top+h;
				totalWidth=s_Skin.Main_padding.left+s_Skin.Main_padding.right+w;
				totalHeight=s_Skin.Main_padding.top+s_Skin.Main_padding.bottom+h;
			}
		}
		else
		{
			m_rContent.left=s_Skin.Submenu_padding.left;
			m_rContent.top=s_Skin.Submenu_padding.top;
			m_rContent.right=s_Skin.Submenu_padding.left+w;
			m_rContent.bottom=s_Skin.Submenu_padding.top+h;
			totalWidth=s_Skin.Submenu_padding.left+s_Skin.Submenu_padding.right+w;
			totalHeight=s_Skin.Submenu_padding.top+s_Skin.Submenu_padding.bottom+h;
		}
	}

	int offs=m_rContent.left, step=maxw;

	for (std::vector<CWindow>::iterator it=m_Toolbars.begin();it!=m_Toolbars.end();++it, offs+=step)
	{
		it->SendMessage(TB_SETBUTTONSIZE,0,MAKELONG(maxw,maxbh));
		it->SetWindowPos(NULL,offs,m_rContent.top,maxw,maxh,SWP_NOZORDER|SWP_NOACTIVATE);
	}

	// create pager
	if (!(m_Options&CONTAINER_MULTICOLUMN) && maxh>maxHeight)
	{
		if (!m_Pager)
			m_Pager=CreateWindow(WC_PAGESCROLLER,L"",WS_CHILD|WS_VISIBLE|PGS_DRAGNDROP|PGS_AUTOSCROLL|PGS_VERT|CCS_NODIVIDER|CCS_NOPARENTALIGN|CCS_NORESIZE,m_rContent.left,m_rContent.top,maxw,maxHeight,m_hWnd,(HMENU)ID_PAGER,g_Instance,NULL);
		else
			m_Pager.SetWindowPos(NULL,m_rContent.left,m_rContent.top,maxw,maxHeight,SWP_NOZORDER);
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

	RECT rc={0,0,totalWidth,totalHeight};
	AdjustWindowRect(&rc,GetWindowLong(GWL_STYLE),FALSE);
	RECT rc0;
	GetWindowRect(&rc0);
	OffsetRect(&rc,(m_Options&CONTAINER_LEFT)?(rc0.left-rc.left):(rc0.right-rc.right),(m_Options&CONTAINER_TOP)?(rc0.top-rc.top):(rc0.bottom-rc.bottom));
	SetWindowPos(NULL,&rc,SWP_NOZORDER|SWP_NOACTIVATE|SWP_NOCOPYBITS|SWP_DEFERERASE);

	// for some reason the region must be set after the call to SetWindowPos. otherwise it doesn't work for RTL windows
	if (m_Region && s_Skin.Main_opacity==MenuSkin::OPACITY_REGION)
	{
		if (SetWindowRgn(m_Region))
			m_Region=NULL; // the OS takes ownership of the region, no need to free
	}

	if (m_Pager.m_hWnd)
	{
		int scroll;
		std::map<unsigned int,int>::iterator it=s_PagerScrolls.find(CalcFNVHash(m_RegName));
		if (it!=s_PagerScrolls.end())
			scroll=it->second; // restore the pager position if the same menu has been opened before
		else if (m_pParent)
			scroll=0; // submenus default to top
		else
		{
			// main menu defaults to bottom
			RECT rc;
			m_Pager.GetClientRect(&rc);
			scroll=rc.bottom;
			m_Toolbars[0].GetWindowRect(&rc);
			scroll=rc.bottom-rc.top-scroll;
		}
		m_Pager.SendMessage(PGM_SETPOS,0,scroll);
	}
}

LRESULT CMenuContainer::OnCreate( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled )
{
	if (!m_pParent)
	{
		if (s_Skin.Main_opacity==MenuSkin::OPACITY_ALPHA)
		{
			MARGINS margins={-1};
			DwmExtendFrameIntoClientArea(m_hWnd,&margins);
		}
		BufferedPaintInit();
	}
	InitToolbars();
	m_ClickTime=GetMessageTime()-10000;
	m_ClickPos.x=m_ClickPos.y=-20;
	m_HotPos=GetMessagePos();
	if (FindSettingBool("EnableAccessibility",true))
	{
		m_pAccessible=new CMenuAccessible(this);
		NotifyWinEvent(EVENT_SYSTEM_MENUPOPUPSTART,m_hWnd,OBJID_CLIENT,CHILDID_SELF);
	}
	else
		m_pAccessible=NULL;
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

		// there is no custom-draw for separators. so here we draw all separators on CDDS_PREPAINT and exclude them from the clip rectangle
		HBITMAP bmpSeparator=NULL;
		bool b32=false;
		int bmpHeight=0;
		const int *slicesX=NULL;
		if (m_pParent)
		{
			bmpSeparator=s_Skin.Submenu_separator;
			b32=s_Skin.Submenu_separator32;
			bmpHeight=s_Skin.Submenu_separatorHeight;
			slicesX=s_Skin.Submenu_separator_slices_X;
		}
		else
		{
			bmpSeparator=s_Skin.Main_separator;
			b32=s_Skin.Main_separator32;
			bmpHeight=s_Skin.Main_separatorHeight;
			slicesX=s_Skin.Main_separator_slices_X;
		}
		if (bmpSeparator)
		{
			// check if there are separators in this toolbar
			int column=0;
			for (int i=0;i<(int)m_Toolbars.size();i++)
			{
				if (m_Toolbars[i].m_hWnd==pDraw->nmcd.hdr.hwndFrom)
				{
					column=i;
					break;
				}
			}
			bool bSep=false;
			for (std::vector<MenuItem>::const_iterator it=m_Items.begin();it!=m_Items.end();++it)
				if (it->id==MENU_SEPARATOR && it->column==column)
				{
					bSep=true;
					break;
				}

			if (bSep)
			{
				// draw the separators for this toolbar and exclude them from the clip rectangle
				HDC hdc=CreateCompatibleDC(pDraw->nmcd.hdc);
				HBITMAP bmp0=(HBITMAP)SelectObject(hdc,bmpSeparator);
				RECT rSrc={0,0,slicesX[0]+slicesX[1]+slicesX[2],bmpHeight};
				RECT rMargins={slicesX[0],bmpHeight,slicesX[2],0};
				CWindow toolbar=pDraw->nmcd.hdr.hwndFrom;
				for (std::vector<MenuItem>::const_iterator it=m_Items.begin();it!=m_Items.end();++it)
				{
					if (it->id!=MENU_SEPARATOR || it->column!=column) continue;
					RECT rc;
					toolbar.SendMessage(TB_GETITEMRECT,it->btnIndex,(LPARAM)&rc);
					if (b32)
						MarginsBlitAlpha(hdc,pDraw->nmcd.hdc,rSrc,rc,rMargins);
					else
						MarginsBlit(hdc,pDraw->nmcd.hdc,rSrc,rc,rMargins);
					ExcludeClipRect(pDraw->nmcd.hdc,rc.left,rc.top,rc.right,rc.bottom);
				}
				SelectObject(hdc,bmp0);
				DeleteDC(hdc);
			}
		}
	}
	else if (pDraw->nmcd.dwDrawStage==CDDS_ITEMPREPAINT)
	{
		res=TBCDRF_USECDCOLORS;
		bool bDisabled=(pDraw->nmcd.uItemState&CDIS_DISABLED) || (pDraw->nmcd.lItemlParam>=ID_OFFSET && m_Items[pDraw->nmcd.lItemlParam-ID_OFFSET].id==MENU_EMPTY);
		if (pDraw->nmcd.uItemState&CDIS_HOT)
		{
			int *slicesX, *slicesY;
			HBITMAP bmp=NULL;
			bool b32=false;
			if (m_pParent)
			{
				// sub-menu
				pDraw->clrText=s_Skin.Submenu_text_color[bDisabled?3:1];
				if (s_Skin.Submenu_selectionColor)
				{
					// set the color for the hot item
					pDraw->clrHighlightHotTrack=s_Skin.Submenu_selection.color;
					res|=TBCDRF_HILITEHOTTRACK;
				}
				else
				{
					bmp=s_Skin.Submenu_selection.bmp;
					b32=s_Skin.Submenu_selection32;
					slicesX=s_Skin.Submenu_selection_slices_X;
					slicesY=s_Skin.Submenu_selection_slices_Y;
				}
			}
			else
			{
				// main menu
				pDraw->clrText=s_Skin.Main_text_color[bDisabled?3:1];
				if (s_Skin.Main_selectionColor)
				{
					// set the color for the hot item
					pDraw->clrHighlightHotTrack=s_Skin.Main_selection.color;
					res|=TBCDRF_HILITEHOTTRACK;
				}
				else
				{
					bmp=s_Skin.Main_selection.bmp;
					b32=s_Skin.Main_selection32;
					slicesX=s_Skin.Main_selection_slices_X;
					slicesY=s_Skin.Main_selection_slices_Y;
				}
			}
			if (bmp)
			{
				HDC hdc=CreateCompatibleDC(pDraw->nmcd.hdc);
				HBITMAP bmp0=(HBITMAP)SelectObject(hdc,bmp);
				RECT rSrc={0,0,slicesX[0]+slicesX[1]+slicesX[2],slicesY[0]+slicesY[1]+slicesY[2]};
				RECT rMargins={slicesX[0],slicesY[0],slicesX[2],slicesY[2]};
				int w=pDraw->nmcd.rc.right-pDraw->nmcd.rc.left;
				int h=pDraw->nmcd.rc.bottom-pDraw->nmcd.rc.top;
				if (rMargins.left>w) rMargins.left=w;
				if (rMargins.right>w) rMargins.right=w;
				if (rMargins.top>h) rMargins.top=h;
				if (rMargins.bottom>h) rMargins.bottom=h;
				if (b32)
					MarginsBlitAlpha(hdc,pDraw->nmcd.hdc,rSrc,pDraw->nmcd.rc,rMargins);
				else
					MarginsBlit(hdc,pDraw->nmcd.hdc,rSrc,pDraw->nmcd.rc,rMargins);
				SelectObject(hdc,bmp0);
				DeleteDC(hdc);
			}
		}
		else
		{
			// set the color for the (Empty) items
			if (m_pParent)
				pDraw->clrText=s_Skin.Submenu_text_color[bDisabled?2:0];
			else
				pDraw->clrText=s_Skin.Main_text_color[bDisabled?2:0];
		}
		res|=CDRF_NOTIFYPOSTPAINT|TBCDRF_NOMARK|TBCDRF_NOOFFSET|TBCDRF_NOEDGES;
		bHandled=TRUE;
	}
	else if (pDraw->nmcd.dwDrawStage==CDDS_ITEMPOSTPAINT && pDraw->nmcd.lItemlParam>=ID_OFFSET && m_Items[pDraw->nmcd.lItemlParam-ID_OFFSET].bFolder)
	{
		// draw a small triangle for the submenus
		bool bDisabled=(pDraw->nmcd.uItemState&CDIS_DISABLED) || (pDraw->nmcd.lItemlParam>=ID_OFFSET && m_Items[pDraw->nmcd.lItemlParam-ID_OFFSET].id==MENU_EMPTY);
		int idx;
		if (pDraw->nmcd.uItemState&CDIS_HOT)
			idx=bDisabled?3:1;
		else
			idx=bDisabled?2:0;

		if (m_pParent)
			SetDCBrushColor(pDraw->nmcd.hdc,s_Skin.Submenu_text_color[idx]);
		else
			SetDCBrushColor(pDraw->nmcd.hdc,s_Skin.Main_text_color[idx]);
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
		SetWindowText(hwndDlg,FindTranslation("Menu.RenameTitle",L"Rename"));
		SetDlgItemText(hwndDlg,IDC_LABEL,FindTranslation("Menu.RenamePrompt",L"&New name:"));
		SetDlgItemText(hwndDlg,IDOK,FindTranslation("Menu.RenameOK",L"OK"));
		SetDlgItemText(hwndDlg,IDCANCEL,FindTranslation("Menu.RenameCancel",L"Cancel"));
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

static void SetShutdownPrivileges( void )
{
	HANDLE hToken;
	if (OpenProcessToken(GetCurrentProcess(),TOKEN_ADJUST_PRIVILEGES|TOKEN_QUERY,&hToken))
	{
		TOKEN_PRIVILEGES tp={1};
		if (LookupPrivilegeValue(NULL,L"SeShutdownPrivilege",&tp.Privileges[0].Luid))
			tp.Privileges[0].Attributes=SE_PRIVILEGE_ENABLED;
		AdjustTokenPrivileges(hToken,FALSE,&tp,sizeof(TOKEN_PRIVILEGES),NULL,NULL); 
		CloseHandle(hToken);
	}
}

static void ExecuteCommand( const wchar_t *command )
{
	wchar_t exe[_MAX_PATH];
	command=GetToken(command,exe,_countof(exe),L" ");
	ShellExecute(NULL,NULL,exe,command,NULL,SW_SHOWNORMAL);
}

// Dialog proc for the Log Off dialog box
static INT_PTR CALLBACK LogOffDlgProc( HWND hwndDlg, UINT uMsg, WPARAM wParam, LPARAM lParam )
{
	if (uMsg==WM_INITDIALOG)
	{
		// translate text
		SendDlgItemMessage(hwndDlg,IDC_STATICICON,STM_SETICON,lParam,0);
		SetWindowText(hwndDlg,FindTranslation("Menu.LogoffTitle",L"Log Off Windows"));
		SetDlgItemText(hwndDlg,IDC_PROMPT,FindTranslation("Menu.LogoffPrompt",L"Are you sure you want to log off?"));
		SetDlgItemText(hwndDlg,IDOK,FindTranslation("Menu.LogoffYes",L"&Log Off"));
		SetDlgItemText(hwndDlg,IDCANCEL,FindTranslation("Menu.LogoffNo",L"&No"));
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
		const StdMenuItem *pSubMenu=item.pStdItem?item.pStdItem->submenu:NULL;

		int options=CONTAINER_DRAG;
		if (item.id==MENU_CONTROLPANEL)
			options|=CONTAINER_NOSUBFOLDERS;
		if (item.id!=MENU_DOCUMENTS && item.id!=MENU_CONTROLPANEL && item.id!=MENU_NETWORK && item.id!=MENU_PRINTERS)
			options|=CONTAINER_DROP;
		if (item.id==MENU_DOCUMENTS)
			options|=CONTAINER_DOCUMENTS|CONTAINER_ADDTOP;
		if (item.bPrograms)
			options|=CONTAINER_PROGRAMS;
		if (item.bLink || (m_Options&CONTAINER_LINK))
			options|=CONTAINER_LINK;

		if (!s_bScrollMenus && !(options&CONTAINER_LINK) && ((m_Options&CONTAINER_MULTICOLUMN) || item.id==MENU_PROGRAMS))
			options|=CONTAINER_MULTICOLUMN;
		CMenuContainer *pMenu=new CMenuContainer(this,index,options,pSubMenu,item.pItem1,item.pItem2,m_RegName+L"\\"+item.name);
		pMenu->InitItems();

		DWORD animFlags;
		{
			const wchar_t *str=FindSetting("SubMenuAnimation");
			if (str && _wcsicmp(str,L"none")==0) animFlags=AW_ACTIVATE;
			else if (str && _wcsicmp(str,L"fade")==0) animFlags=AW_ACTIVATE|AW_BLEND;
			else if (str && _wcsicmp(str,L"slide")==0) animFlags=AW_ACTIVATE|AW_SLIDE;
			else
			{
				DWORD fade;
				SystemParametersInfo(SPI_GETMENUFADE,NULL,&fade,0);
				animFlags=AW_ACTIVATE|(fade?AW_BLEND:AW_SLIDE);
			}
		}

		BOOL animate;
		if ((animFlags&(AW_BLEND|AW_SLIDE))==0 || m_Submenu>=0 || s_bRTL)
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

		pMenu->Create(NULL,NULL,NULL,s_SubmenuStyle,WS_EX_TOOLWINDOW|WS_EX_TOPMOST|(s_bRTL?WS_EX_LAYOUTRTL:0));
		RECT rc2;
		pMenu->GetWindowRect(&rc2);

		RECT border={-s_Skin.Submenu_padding.left,-s_Skin.Submenu_padding.top,s_Skin.Submenu_padding.right,s_Skin.Submenu_padding.bottom};
		AdjustWindowRect(&border,s_SubmenuStyle,FALSE);

		// position new menu
		int w=rc2.right-rc2.left;
		int h=rc2.bottom-rc2.top;

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
				if (!s_bRTL)
				{
					int minx=m_pParent?s_MainRect.left:(rc.right+border.left);
					if (rc2.left<minx)
					{
						rc2.left=minx;
						rc2.right=minx+w;
					}
				}
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
				if (s_bRTL)
				{
					int maxx=m_pParent?s_MainRect.right:(rc.left+border.right);
					if (rc2.right>maxx)
					{
						rc2.left=maxx-w;
						rc2.right=maxx;
					}
				}
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
			int speed=_wtol(FindSetting("SubMenuAnimationSpeed",L""));
			if (speed<=0) speed=MENU_ANIM_SPEED_SUBMENU;
			else if (speed>=10000) speed=10000;
			AnimateWindow(pMenu->m_hWnd,speed,animFlags);
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

		if (!item.pItem1)
		{
			if (item.id<MENU_PROGRAMS) return; // non-executable item
			if (item.pStdItem && item.pStdItem->submenu && !item.pStdItem->command)
				return; // non-executable item
		}

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

	if (!item.pItem1 && !(item.id==MENU_EMPTY && type==ACTIVATE_MENU))
	{
		// regular command item
		if (type!=ACTIVATE_EXECUTE) return;

		if (GetKeyState(VK_SHIFT)<0)
			LockSetForegroundWindow(LSFW_LOCK);
		else
		{
			FadeOutItem(index);
			// flush all messages to close the menus
			// m_hWnd is not valid after this point
			MSG msg;
			while (PeekMessage(&msg,NULL,0,0,PM_REMOVE))
			{
				TranslateMessage(&msg);
				DispatchMessage(&msg);
			}
		}

		// special handling for command items
		CComPtr<IShellDispatch2> pShellDisp;
		switch (item.id)
		{
			case MENU_TASKBAR: // show taskbar properties
				if (SUCCEEDED(CoCreateInstance(CLSID_Shell,NULL,CLSCTX_SERVER,IID_IShellDispatch2,(void**)&pShellDisp)))
					pShellDisp->TrayProperties();
				break;
			case MENU_CLASSIC_SETTINGS: // show our settings
#ifdef ALLOW_DEACTIVATE
				EditSettings(false);
#else
				EditSettings(true);
#endif
				break;
			case MENU_SEARCH_FILES: // show the search UI
				{
					const wchar_t *command=FindSetting("SearchFilesCommand");
					if (command)
					{
						wchar_t buf[1024];
						wcscpy_s(buf,command);
						DoEnvironmentSubst(buf,_countof(buf));
						ExecuteCommand(buf);
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
					SHELLEXECUTEINFO execute={sizeof(execute),SEE_MASK_DOENVSUBST,NULL,L"open"};
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
					if (!FindSettingBool("ConfirmLogOff",false))
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
			case MENU_RESTART: // restart
				SetShutdownPrivileges();
				ExitWindowsEx(EWX_REBOOT,0);
				break;
			case MENU_SWITCHUSER: // switch_user
				WTSDisconnectSession(WTS_CURRENT_SERVER_HANDLE,WTS_CURRENT_SESSION,FALSE); // same as "disconnect"
				break;
			case MENU_SHUTDOWN: // shutdown, don't ask
				SetShutdownPrivileges();
				ExitWindowsEx(EWX_SHUTDOWN,0);
				break;
			case MENU_SLEEP:
				SetSuspendState(FALSE,FALSE,FALSE);
				break;
			case MENU_HIBERNATE:
				SetSuspendState(TRUE,FALSE,FALSE);
				break;
			case MENU_DISCONNECT: // disconnect the current Terminal Services session (remote desktop)
				WTSDisconnectSession(WTS_CURRENT_SERVER_HANDLE,WTS_CURRENT_SESSION,FALSE);
				break;
			case MENU_UNDOCK: // undock the PC
				if (SUCCEEDED(CoCreateInstance(CLSID_Shell,NULL,CLSCTX_SERVER,IID_IShellDispatch2,(void**)&pShellDisp)))
					pShellDisp->EjectPC();
				break;
			case MENU_SHUTDOWN_BOX: // shutdown - ask to shutdown, log off, sleep, etc
				if (SUCCEEDED(CoCreateInstance(CLSID_Shell,NULL,CLSCTX_SERVER,IID_IShellDispatch2,(void**)&pShellDisp)))
					pShellDisp->ShutdownWindows();
				break;
			default:
				if (item.pStdItem && item.pStdItem->command && *item.pStdItem->command)
				{
					wchar_t buf[1024];
					wcscpy_s(buf,item.pStdItem->command);
					DoEnvironmentSubst(buf,_countof(buf));
					ExecuteCommand(buf);
					return;
				}
		}
		return;
	}

	// create a context menu for the selected item. the context menu can be shown (ACTIVATE_MENU) or its default
	// item can be executed automatically (ACTIVATE_EXECUTE)
	CComPtr<IContextMenu> pMenu;
	HMENU menu=CreatePopupMenu();

	CComPtr<IShellFolder> pFolder;
	PCUITEMID_CHILD pidl;

	if (item.id==MENU_EMPTY)
	{
		s_pDesktop->BindToObject(m_Path1,NULL,IID_IShellFolder,(void**)&pFolder);
	}
	else
	{
		SHBindToParent(item.pItem1,IID_IShellFolder,(void**)&pFolder,&pidl);
		if (FAILED(pFolder->GetUIObjectOf(g_OwnerWindow,1,&pidl,IID_IContextMenu,NULL,(void**)&pMenu)))
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

		if (item.bFolder && item.pItem2)
		{
			// context menu for a double folder - remove most command,s add Open All Users
			int n=GetMenuItemCount(menu);
			for (int i=0;i<n;i++)
			{
				int id=GetMenuItemID(menu,i);
				if (id>0)
				{
					char command[256];
					if (FAILED(pMenu->GetCommandString(id-CMD_LAST,GCS_VERBA,NULL,command,_countof(command))))
						command[0]=0;
					if (_stricmp(command,"open")==0)
					{
						InsertMenu(menu,i+1,MF_BYPOSITION|MF_STRING,CMD_OPEN_ALL,FindTranslation("Menu.OpenAll",L"O&pen All Users"));
						InsertMenu(menu,i+2,MF_BYPOSITION|MF_SEPARATOR,0,0);
						i+=2;
						n+=2;
						continue;
					}
					else if (_stricmp(command,"rename")==0 || _stricmp(command,"delete")==0)
					{
						if (item.id!=MENU_PROGRAMS) continue;
					}
					else if (_stricmp(command,"properties")==0)
						continue;
				}
				DeleteMenu(menu,i,MF_BYPOSITION);
				i--;
				n--;
			}
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
		HBITMAP shellBmp=NULL;
		HBITMAP newFolderBmp=NULL;
		if ((item.id==MENU_NO || item.id==MENU_EMPTY) && (m_Options&CONTAINER_DROP))// clicked on a movable item
		{
			AppendMenu(menu,MF_SEPARATOR,0,0);
			HMENU menu2=menu;
			if (FindSettingBool("CascadingMenu",false))
				menu2=CreatePopupMenu();
			bool bSort=false, bNew=false;
			int n=0;
			for (std::vector<MenuItem>::const_iterator it=m_Items.begin();it!=m_Items.end();++it)
				if (it->id==MENU_NO)
					n++;
			if (n>1)
				bSort=true; // more than 1 movable items
			wchar_t path[_MAX_PATH];
			if (pFolder && FindSettingBool("ShowNewFolder",true) && SHGetPathFromIDList(m_Path1,path))
				bNew=true;

			if (bSort)
				AppendMenu(menu2,MF_STRING,CMD_SORT,FindTranslation("Menu.SortByName",L"Sort &by Name"));

			AppendMenu(menu2,MF_STRING,CMD_AUTOSORT,FindTranslation("Menu.AutoArrange",L"&Auto Arrange"));
			if (m_Options&CONTAINER_AUTOSORT)
			{
				EnableMenuItem(menu2,CMD_SORT,MF_BYCOMMAND|MF_GRAYED);
				CheckMenuItem(menu2,CMD_AUTOSORT,MF_BYCOMMAND|MF_CHECKED);
			}

			if (bNew)
				AppendMenu(menu2,MF_STRING,CMD_NEWFOLDER,FindTranslation("Menu.NewFolder",L"New Folder"));

			if (bNew || menu!=menu2)
			{
				int size=16;
				BITMAPINFO bi={0};
				bi.bmiHeader.biSize=sizeof(BITMAPINFOHEADER);
				bi.bmiHeader.biWidth=bi.bmiHeader.biHeight=size;
				bi.bmiHeader.biPlanes=1;
				bi.bmiHeader.biBitCount=32;
				HDC hdc=CreateCompatibleDC(NULL);
				RECT rc={0,0,size,size};

				if (bNew)
				{
					HMODULE hShell32=GetModuleHandle(L"shell32.dll");
					HICON icon=(HICON)LoadImage(hShell32,MAKEINTRESOURCE(319),IMAGE_ICON,size,size,LR_DEFAULTCOLOR);
					if (icon)
					{
						newFolderBmp=CreateDIBSection(hdc,&bi,DIB_RGB_COLORS,NULL,NULL,0);
						HGDIOBJ bmp0=SelectObject(hdc,newFolderBmp);
						FillRect(hdc,&rc,(HBRUSH)GetStockObject(BLACK_BRUSH));
						DrawIconEx(hdc,0,0,icon,size,size,0,NULL,DI_NORMAL);
						SelectObject(hdc,bmp0);
						DeleteObject(icon);

						MENUITEMINFO mii={sizeof(mii)};
						mii.fMask=MIIM_BITMAP;
						mii.hbmpItem=newFolderBmp;
						SetMenuItemInfo(menu2,CMD_NEWFOLDER,FALSE,&mii);
					}
				}
				if (menu!=menu2)
				{
					int idx=GetMenuItemCount(menu);
					AppendMenu(menu,MF_POPUP,(UINT_PTR)menu2,FindTranslation("Menu.Organize",L"Organize Start menu"));
					HICON icon=(HICON)LoadImage(g_Instance,MAKEINTRESOURCE(IDI_APPICON),IMAGE_ICON,size,size,LR_DEFAULTCOLOR);
					if (icon)
					{
						shellBmp=CreateDIBSection(hdc,&bi,DIB_RGB_COLORS,NULL,NULL,0);
						HGDIOBJ bmp0=SelectObject(hdc,shellBmp);
						FillRect(hdc,&rc,(HBRUSH)GetStockObject(BLACK_BRUSH));
						DrawIconEx(hdc,0,0,icon,size,size,0,NULL,DI_NORMAL);
						SelectObject(hdc,bmp0);
						DeleteObject(icon);

						MENUITEMINFO mii={sizeof(mii)};
						mii.fMask=MIIM_BITMAP;
						mii.hbmpItem=shellBmp;
						SetMenuItemInfo(menu,idx,TRUE,&mii);
					}
				}
				DeleteDC(hdc);
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
		if (newFolderBmp) DeleteObject(newFolderBmp);
		if (shellBmp) DeleteObject(shellBmp);
	}

	if (type==ACTIVATE_EXECUTE)
	{
		if (GetKeyState(VK_SHIFT)<0)
			LockSetForegroundWindow(LSFW_LOCK);
		else
			FadeOutItem(index);
	}

	// handle our standard commands
	if (res==CMD_OPEN_ALL)
	{
		SHELLEXECUTEINFO execute={sizeof(execute),SEE_MASK_IDLIST|SEE_MASK_INVOKEIDLIST};
		execute.lpVerb=L"open";
		execute.lpIDList=item.pItem2;
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

	if (res==CMD_AUTOSORT)
	{
		CRegKey regOrder;
		if (regOrder.Open(HKEY_CURRENT_USER,m_RegName)!=ERROR_SUCCESS)
			regOrder.Create(HKEY_CURRENT_USER,m_RegName);
		if (m_Options&CONTAINER_AUTOSORT)
			regOrder.SetBinaryValue(L"Order",NULL,0);
		else
		{
			DWORD cAuto='AUTO';
			regOrder.SetBinaryValue(L"Order",&cAuto,4);
		}
		PostRefreshMessage();
	}

	if (res==CMD_NEWFOLDER)
	{
		s_pDragSource=this; // HACK - prevent the menu from closing
		g_RenameText=item.name;
		g_RenamePos=pPt;
		for (std::vector<CMenuContainer*>::iterator it=s_Menus.begin();it!=s_Menus.end();++it)
			(*it)->EnableWindow(FALSE); // disable all menus

		CComPtr<IContextMenu> pMenu2;
		HMENU menu2=CreatePopupMenu();

		std::vector<unsigned int> items;
		{
			CComPtr<IEnumIDList> pEnum;
			if (pFolder->EnumObjects(NULL,SHCONTF_FOLDERS,&pEnum)!=S_OK) pEnum=NULL;

			PITEMID_CHILD pidl;
			while (pEnum && pEnum->Next(1,&pidl,NULL)==S_OK)
			{
				STRRET str;
				if (SUCCEEDED(pFolder->GetDisplayNameOf(pidl,SHGDN_INFOLDER|SHGDN_FORPARSING,&str)))
				{
					wchar_t *name;
					StrRetToStr(&str,pidl,&name);
					items.push_back(CalcFNVHash(name));
					CoTaskMemFree(name);
				}
				ILFree(pidl);
			}
		}

		if (SUCCEEDED(pFolder->CreateViewObject(g_OwnerWindow,IID_IContextMenu,(void**)&pMenu2)))
		{
			if (SUCCEEDED(pMenu2->QueryContextMenu(menu2,0,1,32767,CMF_NORMAL)))
			{
				CMINVOKECOMMANDINFOEX info={sizeof(info),CMIC_MASK_UNICODE};
				info.lpVerb="NewFolder";
				info.lpVerbW=L"NewFolder";
				info.nShow=SW_SHOWNORMAL;
				info.fMask|=CMIC_MASK_NOASYNC;
				info.hwnd=g_OwnerWindow;
				pMenu2->InvokeCommand((CMINVOKECOMMANDINFO*)&info);
			}
		}
		DestroyMenu(menu2);

		PITEMID_CHILD newPidl=NULL;
		unsigned int newHash=0;
		{
			CComPtr<IEnumIDList> pEnum;
			if (pFolder->EnumObjects(NULL,SHCONTF_FOLDERS,&pEnum)!=S_OK) pEnum=NULL;

			PITEMID_CHILD pidl;
			while (pEnum && pEnum->Next(1,&pidl,NULL)==S_OK)
			{
				STRRET str;
				if (SUCCEEDED(pFolder->GetDisplayNameOf(pidl,SHGDN_INFOLDER|SHGDN_FORPARSING,&str)))
				{
					wchar_t *name;
					StrRetToStr(&str,pidl,&name);
					unsigned int hash=CalcFNVHash(name);
					if (std::find(items.begin(),items.end(),hash)==items.end())
					{
						if (SUCCEEDED(pFolder->GetDisplayNameOf(pidl,SHGDN_INFOLDER|SHGDN_FOREDITING,&str)))
						{
							wchar_t *name2;
							StrRetToStr(&str,pidl,&name2);
							g_RenameText=name2;
							CoTaskMemFree(name2);
							StrRetToStr(&str,pidl,&name2);
						}
						else
							g_RenameText=name;
						CharUpper(name);
						newHash=CalcFNVHash(name);
						CoTaskMemFree(name);
						newPidl=pidl;
						break;
					}
					CoTaskMemFree(name);
				}
				ILFree(pidl);
			}
		}

		PostRefreshMessage();
		PostMessage(MCM_SETHOTITEM,newHash);
		// show the Rename dialog box
		if (newPidl && DialogBox(g_Instance,MAKEINTRESOURCE(s_bRTL?IDD_RENAMER:IDD_RENAME),g_OwnerWindow,RenameDlgProc))
		{
			pFolder->SetNameOf(g_OwnerWindow,newPidl,g_RenameText,SHGDN_INFOLDER|SHGDN_FOREDITING,NULL);
			PostRefreshMessage();
		}
		for (std::vector<CMenuContainer*>::iterator it=s_Menus.begin();it!=s_Menus.end();++it)
			(*it)->EnableWindow(TRUE); // enable all menus
		SetForegroundWindow(m_hWnd);
		SetActiveWindow();
		m_Toolbars[0].SetFocus();
		s_pDragSource=NULL;
		DestroyMenu(menu);
		return;
	}

	// handle the shell commands
	if (res>=CMD_LAST)
	{
		// handle special verbs
		char command[256];
		if (FAILED(pMenu->GetCommandString(res-CMD_LAST,GCS_VERBA,NULL,command,_countof(command))))
			command[0]=0;
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

			bool bRenamed=DialogBox(g_Instance,MAKEINTRESOURCE(s_bRTL?IDD_RENAMER:IDD_RENAME),g_OwnerWindow,RenameDlgProc)!=0;

			if (bRenamed)
			{
				// perform the rename operation
				PITEMID_CHILD newPidl;
				if (SUCCEEDED(pFolder->SetNameOf(g_OwnerWindow,pidl,g_RenameText,SHGDN_INFOLDER|SHGDN_FOREDITING,&newPidl)))
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

						if (!(m_Options&CONTAINER_AUTOSORT))
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
					ILFree(newPidl);
				}
				PostRefreshMessage();
			}
			for (std::vector<CMenuContainer*>::iterator it=s_Menus.begin();it!=s_Menus.end();++it)
				(*it)->EnableWindow(TRUE); // enable all menus
			SetForegroundWindow(m_hWnd);
			SetActiveWindow();
			m_Toolbars[0].SetFocus();
			s_pDragSource=NULL;
			DestroyMenu(menu);
			return;
		}

		bool bRefresh=(_stricmp(command,"delete")==0 || _stricmp(command,"link")==0);

		CMINVOKECOMMANDINFOEX info={sizeof(info),CMIC_MASK_UNICODE};
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

		if (bRefresh)
			info.fMask|=CMIC_MASK_NOASYNC; // wait for delete/link commands to finish so we can refresh the menu

		s_pDragSource=this; // prevent the menu from closing. the command may need a HWND to show its UI
		for (std::vector<CMenuContainer*>::iterator it=s_Menus.begin();it!=s_Menus.end();++it)
			(*it)->EnableWindow(FALSE);
		info.hwnd=g_OwnerWindow;

		::SetForegroundWindow(g_OwnerWindow);
		RECT rc;
		GetWindowRect(&rc);
		::SetWindowPos(g_OwnerWindow,HWND_TOPMOST,rc.left,rc.top,rc.right-rc.left,rc.bottom-rc.top,0);
		HRESULT hr=pMenu->InvokeCommand((LPCMINVOKECOMMANDINFO)&info);

		for (std::vector<CMenuContainer*>::iterator it=s_Menus.begin();it!=s_Menus.end();++it)
			if (!(*it)->m_bDestroyed)
				(*it)->EnableWindow(TRUE);
		if (bRefresh && !m_bDestroyed)
		{
			SetForegroundWindow(m_hWnd);
			SetActiveWindow();
			m_Toolbars[0].SetFocus();
		}
		s_pDragSource=NULL;

		if (bRefresh)
			PostRefreshMessage(); // refresh the menu after an item was deleted or created
	}
	DestroyMenu(menu);
}

CMenuFader::CMenuFader( HBITMAP bmp, HRGN region, int duration, RECT &rect )
{
	m_Bitmap=bmp;
	m_Region=region;
	m_Duration=duration;
	m_Rect=rect;
	s_Faders.push_back(this);
}

CMenuFader::~CMenuFader( void )
{
	if (m_Bitmap) DeleteObject(m_Bitmap);
	if (m_Region) DeleteObject(m_Region);
	s_Faders.erase(std::find(s_Faders.begin(),s_Faders.end(),this));
}

void CMenuFader::Create( void )
{
	CWindowImpl<CMenuFader>::Create(NULL,&m_Rect,NULL,WS_POPUP,WS_EX_TOOLWINDOW|WS_EX_TOPMOST|WS_EX_LAYERED);
	ShowWindow(SW_SHOWNOACTIVATE);
	if (m_Region)
	{
		SetWindowRgn(m_Region);
		m_Region=NULL;
	}
	SetTimer(1,20);
	m_Time0=GetMessageTime();
	m_LastTime=0;
	PostMessage(WM_TIMER,0,0);
	SetLayeredWindowAttributes(m_hWnd,0,255,LWA_ALPHA);
}

LRESULT CMenuFader::OnEraseBkgnd( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled )
{
	RECT rc;
	GetClientRect(&rc);
	HDC hdc=(HDC)wParam;

	// draw the background
	HDC hdc2=CreateCompatibleDC(hdc);
	HGDIOBJ bmp0=SelectObject(hdc2,m_Bitmap);
	BitBlt(hdc,0,0,rc.right,rc.bottom,hdc2,0,0,SRCCOPY);
	SelectObject(hdc2,bmp0);
	DeleteDC(hdc2);
	return 1;
}

LRESULT CMenuFader::OnTimer( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled )
{
	int t=GetMessageTime()-m_Time0;
	const int MAX_DELTA=80; // allow at most 80ms between redraws. if more, slow down time
	if (t>MAX_DELTA+m_LastTime)
	{
		m_Time0+=t-MAX_DELTA-m_LastTime;
		t=MAX_DELTA+m_LastTime;
	}
	m_LastTime=t;
	if (t<m_Duration)
	{
		SetLayeredWindowAttributes(m_hWnd,0,(m_Duration-t)*255/m_Duration,LWA_ALPHA);
		RedrawWindow();
	}
	else
	{
		KillTimer(1);
		PostMessage(WM_CLOSE);
	}
	return 0;
}

void CMenuFader::ClearAll( void )
{
	while (!s_Faders.empty())
		s_Faders[0]->SendMessage(WM_CLOSE);
}

static DWORD WINAPI FaderThreadProc( void *param )
{
	((CMenuFader*)param)->Create();
	MSG msg;
	while (GetMessage(&msg,NULL,0,0))
	{
		TranslateMessage(&msg);
		DispatchMessage(&msg);
	}
	return 0;
}

void CMenuContainer::FadeOutItem( int index )
{
	if (s_bRTL) return; // fading doesn't work with RTL because of WM_PRINTCLIENT problems
	int speed=MENU_FADE_SPEED;
	const wchar_t *str=FindSetting("MenuFadeSpeed");
	if (!str)
	{
		DWORD fade;
		SystemParametersInfo(SPI_GETSELECTIONFADE,NULL,&fade,0);
		if (!fade) return;
	}
	else
	{
		speed=_wtol(str);
		if (speed<0) speed=0;
		if (speed>10000) speed=10000;
	}
	if (!speed) return;

	const MenuItem &item=m_Items[index];
	CWindow toolbar=m_Toolbars[item.column];
	RECT rc;
	toolbar.SendMessage(TB_GETITEMRECT,item.btnIndex,(LPARAM)&rc);

	BITMAPINFO dib={sizeof(dib)};
	dib.bmiHeader.biWidth=rc.right-rc.left;
	dib.bmiHeader.biHeight=rc.bottom-rc.top;
	dib.bmiHeader.biPlanes=1;
	dib.bmiHeader.biBitCount=32;
	dib.bmiHeader.biCompression=BI_RGB;

	HDC hdc=CreateCompatibleDC(NULL);
	unsigned int *bits;
	HBITMAP bmp=CreateDIBSection(hdc,&dib,DIB_RGB_COLORS,(void**)&bits,NULL,0);
	HBITMAP bmp0=(HBITMAP)SelectObject(hdc,bmp);
	SetViewportOrgEx(hdc,-rc.left,-rc.top,NULL);

	int *slicesX, *slicesY;
	HBITMAP bmpSel=NULL;
	bool b32=false;
	if (m_pParent)
	{
		// sub-menu
		if (!s_Skin.Submenu_selectionColor)
		{
			bmpSel=s_Skin.Submenu_selection.bmp;
			b32=s_Skin.Submenu_selection32;
			slicesX=s_Skin.Submenu_selection_slices_X;
			slicesY=s_Skin.Submenu_selection_slices_Y;
		}
	}
	else
	{
		// main menu
		if (!s_Skin.Main_selectionColor)
		{
			bmpSel=s_Skin.Main_selection.bmp;
			b32=s_Skin.Main_selection32;
			slicesX=s_Skin.Main_selection_slices_X;
			slicesY=s_Skin.Main_selection_slices_Y;
		}
	}
	HRGN region=NULL;
	if (bmpSel && b32)
	{
		HDC hdc2=CreateCompatibleDC(hdc);
		HBITMAP bmp02=(HBITMAP)SelectObject(hdc2,bmpSel);
		FillRect(hdc,&rc,(HBRUSH)GetStockObject(WHITE_BRUSH));
		RECT rSrc={0,0,slicesX[0]+slicesX[1]+slicesX[2],slicesY[0]+slicesY[1]+slicesY[2]};
		RECT rMargins={slicesX[0],slicesY[0],slicesX[2],slicesY[2]};
		int w=dib.bmiHeader.biWidth;
		int h=dib.bmiHeader.biHeight;
		if (rMargins.left>w) rMargins.left=w;
		if (rMargins.right>w) rMargins.right=w;
		if (rMargins.top>h) rMargins.top=h;
		if (rMargins.bottom>h) rMargins.bottom=h;
		MarginsBlit(hdc2,hdc,rSrc,rc,rMargins);
		SelectObject(hdc2,bmp02);
		DeleteDC(hdc2);
		SelectObject(hdc,bmp0);

		for (int y=0;y<h;y++)
		{
			int minx=-1, maxx=-1;
			int yw=y*w;
			for (int x=0;x<w;x++)
			{
				if ((bits[yw+x]>>24)>=32)
				{
					if (minx==-1) minx=x; // first non-transparent pixel
					if (maxx<x) maxx=x; // last non-transparent pixel
				}
			}
			if (minx>=0)
			{
				maxx++;
				HRGN r=CreateRectRgn(minx,y,maxx,y+1);
				if (!region)
					region=r;
				else
				{
					CombineRgn(region,region,r,RGN_OR);
					DeleteObject(r);
				}
			}
		}

		SelectObject(hdc,bmp);
	}

	SetDCBrushColor(hdc,m_pParent?s_Skin.Submenu_background:s_Skin.Main_background);
	FillRect(hdc,&rc,(HBRUSH)GetStockObject(DC_BRUSH));
	toolbar.SendMessage(WM_PRINTCLIENT,(WPARAM)hdc,PRF_CLIENT);

	SelectObject(hdc,bmp0);
	DeleteDC(hdc);

	toolbar.ClientToScreen(&rc);
	CMenuFader *pFader=new CMenuFader(bmp,region,speed,rc);
	CreateThread(NULL,0,FaderThreadProc,pFader,0,NULL);
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
		// skip separators and the (Empty) item in the top menu if hidden
		while (m_Items[index].id==MENU_SEPARATOR || (!m_pParent && !s_bShowTopEmpty && m_Items[index].id==MENU_EMPTY))
			index=(index+n-1)%n;
		index=(index+n-1)%n;
		while (m_Items[index].id==MENU_SEPARATOR || (!m_pParent && !s_bShowTopEmpty && m_Items[index].id==MENU_EMPTY))
			index=(index+n-1)%n;
		ActivateItem(index,ACTIVATE_SELECT,NULL);
	}
	if (pKey->nVKey==VK_DOWN)
	{
		// next item
		while (index>=0 && (m_Items[index].id==MENU_SEPARATOR || (!m_pParent && !s_bShowTopEmpty && m_Items[index].id==MENU_EMPTY)))
			index=(index+1)%n;
		index=(index+1)%n;
		while (m_Items[index].id==MENU_SEPARATOR || (!m_pParent && !s_bShowTopEmpty && m_Items[index].id==MENU_EMPTY))
			index=(index+1)%n;
		ActivateItem(index,ACTIVATE_SELECT,NULL);
	}
	if (pKey->nVKey==VK_ESCAPE || (pKey->nVKey==VK_LEFT && !s_bRTL) || (pKey->nVKey==VK_RIGHT && s_bRTL))
	{
		// close top menu
		if (!s_Menus[s_Menus.size()-1]->m_bDestroyed)
			s_Menus[s_Menus.size()-1]->PostMessage(WM_CLOSE);
		if (s_Menus.size()>=2 && !s_Menus[s_Menus.size()-2]->m_bDestroyed)
			s_Menus[s_Menus.size()-2]->SetActiveWindow();
		if (s_Menus.size()==1)
		{
			// HACK: stops the call to SetActiveWindow(NULL). The correct behavior is to not close the taskbar when Esc is pressed
			s_TaskbarState&=~ABS_AUTOHIDE;
		}
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

LRESULT CMenuContainer::OnSetHotItem( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled )
{
	for (int i=0;i<(int)m_Items.size();i++)
	{
		if (m_Items[i].nameHash==wParam)
		{
			ActivateItem(i,ACTIVATE_SELECT,NULL);
			break;
		}
	}
	return 0;
}

LRESULT CMenuContainer::OnHotItemChange( int idCtrl, LPNMHDR pnmh, BOOL& bHandled )
{
	NMTBHOTITEM *pItem=(NMTBHOTITEM*)pnmh;
	if (m_ContextItem!=-1 && pItem->idNew!=m_ContextItem+ID_OFFSET)
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
	if (m_Submenu>=0)
	{
		// don't show infotip if there is a submenu
		return 0;

/*
		an attempt to show the infotip for a folder item even after the submenu is opened. this doesn't work for 2 reasons.
		1) the tooltip and the submenu fight who to be on top. it is not pretty
		2) the toolbar can't decide where to show the infotip when it there is a submenu. sometimes it is next to the mouse, sometimes it is to the right of the toolbar item
		when enabling this code also uncomment the TTS_ALWAYSTIP code in CreateToolbar to get the tip to show even for inactive windows

		int n=(int)s_Menus.size()-1;
		if (s_Menus[n]!=this) // if this is not the top menu
		{
			if (s_Menus[n-1]!=this)
				return 0; // there are at least 2 more menus on top
		}
		// check if the top menu has a hot item
		for (std::vector<CWindow>::iterator it=s_Menus[n]->m_Toolbars.begin();it!=s_Menus[n]->m_Toolbars.end();++it)
		{
			int hot=(int)it->SendMessage(TB_GETHOTITEM);
			if (hot>=0) return 0;
		}*/
	}

	NMTBGETINFOTIP *pTip=(NMTBGETINFOTIP*)pnmh;
	const MenuItem &item=m_Items[pTip->lParam-ID_OFFSET];
	if (item.pStdItem && (item.pStdItem->tipKey || item.pStdItem->tip))
	{
		// show the tip for the standard item
		wcscpy_s(pTip->pszText,pTip->cchTextMax,item.pStdItem->tipKey?FindTranslation(item.pStdItem->tipKey,item.pStdItem->tip):item.pStdItem->tip);
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
	else
	{
		bool bFound=false;
		for (std::vector<CWindow>::const_iterator it=m_Toolbars.begin();it!=m_Toolbars.end();++it)
			if (it->m_hWnd==toolbar.m_hWnd)
			{
				bFound=true;
				break;
			}
		if (!bFound) return 0;
	};
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
	if (m_pAccessible)
	{
		NotifyWinEvent(EVENT_SYSTEM_MENUPOPUPEND,m_hWnd,OBJID_CLIENT,CHILDID_SELF);
		m_pAccessible->Reset();
		m_pAccessible->Release();
	}
	// remember the scroll position
	unsigned int key=CalcFNVHash(m_RegName);
	if (m_Pager.m_hWnd)
		s_PagerScrolls[key]=(int)m_Pager.SendMessage(PGM_GETPOS);
	else
		s_PagerScrolls.erase(key);

	m_bDestroyed=true;
	if (!m_pParent)
	{
		// cleanup when the last menu is closed
		EnableStartTooltip(true);
		BufferedPaintUnInit();
		FreeMenuSkin(s_Skin);
		if (s_TaskbarState&ABS_AUTOHIDE)
			::SetActiveWindow(NULL); // close the taskbar if it is auto-hide
	}
	return 0;
}

LRESULT CMenuContainer::OnRefresh( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled )
{
	// updates the menu after drag/drop, delete, or rename operation
	m_bRefreshPosted=false;
	unsigned int key=CalcFNVHash(m_RegName);
	if (m_Pager.m_hWnd)
		s_PagerScrolls[key]=(int)m_Pager.SendMessage(PGM_GETPOS);
	else
		s_PagerScrolls.erase(key);
	InitItems();
	InitToolbars();
	m_Toolbars[0].SetFocus();
	m_Toolbars[0].SendMessage(TB_SETHOTITEM,-1);
	return 0;
}

LRESULT CMenuContainer::OnEraseBkgnd( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled )
{
	// draw the background (bitmap or solid color)
	RECT rc;
	GetClientRect(&rc);
	HDC hdc=(HDC)wParam;

	if (m_Bitmap)
	{
		if (s_Skin.Main_opacity==MenuSkin::OPACITY_GLASS)
		{
			DWM_BLURBEHIND blur={DWM_BB_ENABLE|DWM_BB_BLURREGION,TRUE,m_Region,FALSE};
			DwmEnableBlurBehindWindow(m_hWnd,&blur);
		}

		// draw the background
		HDC hdc2=CreateCompatibleDC(hdc);
		HGDIOBJ bmp0=SelectObject(hdc2,m_Bitmap);
		BitBlt(hdc,0,0,rc.right,rc.bottom,hdc2,0,0,SRCCOPY);
		SelectObject(hdc2,bmp0);
		DeleteDC(hdc2);
	}
	else
	{
		SetDCBrushColor(hdc,m_pParent?s_Skin.Submenu_background:s_Skin.Main_background);
		FillRect(hdc,&rc,(HBRUSH)GetStockObject(DC_BRUSH));
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

	if (g_OwnerWindow && (HWND)lParam==g_OwnerWindow)
		return 0;

	for (std::vector<CMenuContainer*>::reverse_iterator it=s_Menus.rbegin();it!=s_Menus.rend();++it)
		if (!(*it)->m_bDestroyed)
			(*it)->PostMessage(WM_CLOSE);
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
			}
		}
		if (hashes.size()==1 && hashes[0]=='AUTO')
		{
			m_Options|=CONTAINER_AUTOSORT;
			for (std::vector<MenuItem>::iterator it=m_Items.begin();it!=m_Items.end();++it)
				it->btnIndex=0;
		}
		else
		{
			m_Options&=~CONTAINER_AUTOSORT;

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
		}
	}
	else
	{
		for (std::vector<MenuItem>::iterator it=m_Items.begin();it!=m_Items.end();++it)
			it->btnIndex=0;
	}

	// sort by btnIndex, then by bFolder, then by name
	std::sort(m_Items.begin(),m_Items.end());
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

// Toggles the start menu
HWND CMenuContainer::ToggleStartMenu( HWND startButton, bool bKeyboard )
{
	if (!bKeyboard) s_LastFGWindow=NULL;
	if (CloseStartMenu())
		return NULL;

	s_LastFGWindow=GetForegroundWindow();
	SetForegroundWindow(g_StartButton);

	EnableStartTooltip(false);

	// initialize all settings
	StartMenuSettings settings;
	ReadSettings(settings);

	if (!LoadMenuSkin(settings.SkinName,s_Skin,settings.SkinVariation,false))
		LoadDefaultMenuSkin(s_Skin,NULL,false);

	s_bScrollMenus=(settings.ScrollMenus!=0);
	s_bExpandLinks=(settings.ExpandLinks!=0);
	s_MaxRecentDocuments=_wtol(FindSetting("MaxRecentDocuments",L"15"));
	s_ShellFormat=RegisterClipboardFormat(CFSTR_SHELLIDLIST);

	bool bRemote=GetSystemMetrics(SM_REMOTESESSION)!=0;
	wchar_t wabPath[_MAX_PATH]=L"%ProgramFiles%\\Windows Mail\\wab.exe";
	DoEnvironmentSubst(wabPath,_countof(wabPath));
	HANDLE hWab=CreateFile(wabPath,0,FILE_SHARE_WRITE,NULL,OPEN_EXISTING,FILE_ATTRIBUTE_NORMAL,NULL);
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

	if (s_TaskbarState&ABS_AUTOHIDE)
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
			case MENU_SHUTDOWN_BOX:
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
				g_StdOptions[i].options=FindSettingBool("ShowHelp",!SHRestricted(REST_NOSMHELP))?MENU_ENABLED:0;
				break;
			case MENU_RUN:
				g_StdOptions[i].options=FindSettingBool("ShowRun",!SHRestricted(REST_NORUN))?MENU_ENABLED:0;
				break;
			case MENU_TASKBAR:
				g_StdOptions[i].options=!SHRestricted(REST_NOSETTASKBAR)?MENU_ENABLED:0;
				break;
			case MENU_FEATURES:
				g_StdOptions[i].options=(!bNoSetFolders && !SHRestricted(REST_NOCONTROLPANEL))?MENU_ENABLED:0;
				break;
			case MENU_SEARCH:
				g_StdOptions[i].options=FindSettingBool("ShowSearch",!SHRestricted(REST_NOFIND))?MENU_ENABLED:0;
				break;
			case MENU_USERFILES:
				{
					const wchar_t *str=FindSetting("ShowUserFiles");
					if (!str)
						g_StdOptions[i].options=!SHRestricted(REST_NOSMMYDOCS)?MENU_ENABLED:0;
					else
					{
						int x=_wtol(str);
						if (x==1)
							g_StdOptions[i].options=MENU_ENABLED;
						else if (x==2)
							g_StdOptions[i].options=MENU_ENABLED|MENU_EXPANDED;
						else
							g_StdOptions[i].options=0;
					}
				}
				break;
			case MENU_USERDOCUMENTS:
				{
					const wchar_t *str=FindSetting("ShowUserDocuments");
					if (!str)
						g_StdOptions[i].options=!SHRestricted(REST_NOSMMYDOCS)?MENU_ENABLED:0;
					else
					{
						int x=_wtol(str);
						if (x==1)
							g_StdOptions[i].options=MENU_ENABLED;
						else if (x==2)
							g_StdOptions[i].options=MENU_ENABLED|MENU_EXPANDED;
						else
							g_StdOptions[i].options=0;
					}
				}
				break;
			case MENU_USERPICTURES:
				{
					const wchar_t *str=FindSetting("ShowUserPictures");
					if (!str)
						g_StdOptions[i].options=!SHRestricted(REST_NOSMMYPICS)?MENU_ENABLED:0;
					else
					{
						int x=_wtol(str);
						if (x==1)
							g_StdOptions[i].options=MENU_ENABLED;
						else if (x==2)
							g_StdOptions[i].options=MENU_ENABLED|MENU_EXPANDED;
						else
							g_StdOptions[i].options=0;
					}
				}
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
	{
		const wchar_t *str=FindSetting("MenuDelay");
		if (str)
			s_HoverTime=_wtol(str);
		else
			SystemParametersInfo(SPI_GETMENUSHOWDELAY,NULL,&s_HoverTime,0);
	}

	// create the top menu from the Start Menu folders
	PIDLIST_ABSOLUTE path1;
	PIDLIST_ABSOLUTE path2;
	SHGetKnownFolderIDList(FOLDERID_StartMenu,0,NULL,&path1);
	SHGetKnownFolderIDList(FOLDERID_CommonStartMenu,0,NULL,&path2);

	int options=CONTAINER_PROGRAMS|CONTAINER_DRAG|CONTAINER_DROP;
	if (s_Skin.Main_large_icons)
		options|=CONTAINER_LARGE;

	DWORD animFlags;
	{
		const wchar_t *str=FindSetting("MainMenuAnimation");
		if (str && _wcsicmp(str,L"none")==0) animFlags=0;
		else if (str && _wcsicmp(str,L"fade")==0) animFlags=AW_BLEND;
		else if (str && _wcsicmp(str,L"slide")==0) animFlags=AW_SLIDE;
		else
		{
			DWORD fade;
			SystemParametersInfo(SPI_GETMENUFADE,NULL,&fade,0);
			animFlags=(fade?AW_BLEND:AW_SLIDE);
		}
	}

	s_bBehindTaskbar=true;
	s_bShowTopEmpty=false;
	DWORD dwStyle=WS_POPUP;
	s_SubmenuStyle=WS_POPUP;
	bool bTheme=IsAppThemed()!=FALSE;
	if (bTheme)
	{
		BOOL comp;
		if (s_Skin.Main_opacity==MenuSkin::OPACITY_SOLID)
			dwStyle|=WS_BORDER;
		else if (FAILED(DwmIsCompositionEnabled(&comp)) || !comp)
			s_Skin.Main_opacity=MenuSkin::OPACITY_REGION;
		s_SubmenuStyle|=WS_BORDER;
	}
	else
	{
		if (s_Skin.Main_opacity==MenuSkin::OPACITY_SOLID)
			dwStyle|=s_Skin.Main_thin_frame?WS_BORDER:WS_DLGFRAME;
		else
			s_Skin.Main_opacity=MenuSkin::OPACITY_REGION;
		s_SubmenuStyle|=s_Skin.Submenu_thin_frame?WS_BORDER:WS_DLGFRAME;
	}

	RECT margin={0,0,0,0};
	AdjustWindowRect(&margin,s_SubmenuStyle,FALSE);
	s_Skin.Submenu_padding.left+=margin.left; if (s_Skin.Submenu_padding.left<0) s_Skin.Submenu_padding.left=0;
	s_Skin.Submenu_padding.right-=margin.right; if (s_Skin.Submenu_padding.right<0) s_Skin.Submenu_padding.right=0;
	s_Skin.Submenu_padding.top+=margin.top; if (s_Skin.Submenu_padding.top<0) s_Skin.Submenu_padding.top=0;
	s_Skin.Submenu_padding.bottom-=margin.bottom; if (s_Skin.Submenu_padding.bottom<0) s_Skin.Submenu_padding.bottom=0;

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
	}
	s_Skin.Main_padding.right-=margin.right; if (s_Skin.Main_padding.right<0) s_Skin.Main_padding.right=0;
	s_Skin.Main_padding.top+=margin.top; if (s_Skin.Main_padding.top<0) s_Skin.Main_padding.top=0;
	s_Skin.Main_padding.bottom-=margin.bottom; if (s_Skin.Main_padding.bottom<0) s_Skin.Main_padding.bottom=0;

	if (!bTheme)
		memset(&margin,0,sizeof(margin)); // in Classic mode don't offset the main menu by the border size
	RECT rc;
	if ((appbar.uEdge==ABE_LEFT || appbar.uEdge==ABE_RIGHT) && FindSettingBool("ShowNextToTaskbar",false))
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

	CMenuContainer *pStartMenu=new CMenuContainer(NULL,-1,options,ParseCustomMenu(),path1,path2,L"Software\\IvoSoft\\ClassicStartMenu\\Order");
	pStartMenu->InitItems();
	ILFree(path1);
	ILFree(path2);

	bool bTopMost=(s_TaskbarState&ABS_ALWAYSONTOP)!=0;

	if (!pStartMenu->Create(NULL,&rc,NULL,dwStyle,WS_EX_TOOLWINDOW|(bTopMost?WS_EX_TOPMOST:0)|(s_bRTL?WS_EX_LAYOUTRTL:0)))
	{
		FreeMenuSkin(s_Skin);
		return NULL;
	}
	pStartMenu->m_Toolbars[0].SendMessage(TB_SETHOTITEM,-1);

	BOOL animate;
	// toolbars don't handle WM_PRINT correctly with RTL, so AnimateWindow doesn't work. disable animations with RTL
	// AnimateWindow also doesn't work correctly for region, alpha or glass windows
	// so only LTR and solid windows can animate. bummer.
	if ((animFlags&(AW_BLEND|AW_SLIDE))==0 || s_bRTL || s_Skin.Main_opacity!=MenuSkin::OPACITY_SOLID)
		animate=FALSE;
	else
		SystemParametersInfo(SPI_GETMENUANIMATION,NULL,&animate,0);

	if (s_bBehindTaskbar)
		::SetWindowPos(startButton,bTopMost?HWND_TOPMOST:HWND_TOP,0,0,0,0,SWP_NOMOVE|SWP_NOSIZE); // bring the start button on top
	if (animate)
	{
		int speed=_wtol(FindSetting("MainMenuAnimationSpeed",L""));
		if (speed<=0) speed=MENU_ANIM_SPEED;
		else if (speed>=10000) speed=10000;
		AnimateWindow(pStartMenu->m_hWnd,speed,animFlags);
	}
	else
		pStartMenu->ShowWindow(SW_SHOW);
	pStartMenu->m_Toolbars[0].SetFocus();
	pStartMenu->m_Toolbars[0].SendMessage(TB_SETHOTITEM,-1);
	// position the start button on top
	if (s_bBehindTaskbar)
		::SetWindowPos(startButton,bTopMost?HWND_TOPMOST:HWND_TOP,0,0,0,0,SWP_NOMOVE|SWP_NOSIZE|SWP_NOACTIVATE);

	SetForegroundWindow(pStartMenu->m_hWnd);

	return pStartMenu->m_hWnd;
}
