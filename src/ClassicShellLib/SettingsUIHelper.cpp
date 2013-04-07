// Classic Shell (c) 2009-2013, Ivo Beltchev
// The sources for Classic Shell are distributed under the MIT open source license

#define STRICT_TYPED_ITEMIDS
#include <windows.h>
#include <atlbase.h>
#include <atlwin.h>
#include <atlstr.h>
#include "resource.h"
#include "Settings.h"
#include "SettingsUIHelper.h"
#include "ResourceHelper.h"
#include "FNVHash.h"
#include "StringUtils.h"
#include "Translations.h"
#include <shlobj.h>
#include <uxtheme.h>
#include <wininet.h>
#include <map>
#include <algorithm>

const GUID FOLDERID_DesktopRoot={'DESK', 'TO', 'P', {'D', 'E', 'S', 'K', 'T', 'O', 'P', 0x00}};

static bool IsVariantTrue( const CComVariant &var )
{
	return (var.vt==VT_I4 && var.intVal==1);
}

///////////////////////////////////////////////////////////////////////////////

void CTreeItem::SetCommand( CString command, const CStdCommand *pStdCommands )
{
	this->command=command;
	pStdCommand=NULL;
	for (int i=0;pStdCommands[i].name;i++)
		if (_wcsicmp(pStdCommands[i].name,command)==0)
			pStdCommand=&pStdCommands[i];
}

unsigned int CTreeItem::GetIconKey( void ) const
{
	if (!icon.IsEmpty())
		return CalcFNVHash(icon);
	else if (pStdCommand && pStdCommand->knownFolder)
		return CalcFNVHash(pStdCommand->knownFolder,sizeof(GUID));
	else if (!link.IsEmpty())
		return CalcFNVHash(link);
	return 0;
}

HICON CTreeItem::LoadIcon( bool bSmall, std::vector<HMODULE> &modules ) const
{
	if (!icon.IsEmpty())
	{
		if (_wcsicmp(icon,L"none")==0)
			return bSmall?HICON_NONE:NULL;
		return ::LoadIcon(GetSystemMetrics(bSmall?SM_CXSMICON:SM_CXICON),icon,modules);
	}
	else
	{
		PIDLIST_ABSOLUTE pidl=NULL;
		if (pStdCommand && pStdCommand->knownFolder)
		{
			if (FAILED(SHGetKnownFolderIDList(*pStdCommand->knownFolder,0,NULL,&pidl)))
				pidl=NULL;
		}
		else if (!link.IsEmpty())
		{
			const wchar_t *c=wcschr(link,'|');
			if (c)
			{
				for (c++;*c==' ';)
					c++;
			}
			else
				c=link;
			wchar_t buf[_MAX_PATH];
			Strcpy(buf,_countof(buf),c);
			DoEnvironmentSubst(buf,_countof(buf));

			ShParseDisplayName(buf,&pidl,0,NULL);
		}
		if (pidl)
		{
			HICON hIcon=::LoadIcon(GetSystemMetrics(bSmall?SM_CXSMICON:SM_CXICON),pidl);
			ILFree(pidl);
			return hIcon;
		}
	}
	return NULL;
}

unsigned int CTreeItem::GetIconDKey( unsigned int iconKey ) const
{
	if (iconD.IsEmpty())
		return ~iconKey;
	return CalcFNVHash(iconD);
}

HICON CTreeItem::LoadIconD( HICON hIcon, std::vector<HMODULE> &modules ) const
{
	if (!iconD.IsEmpty())
	{
		if (_wcsicmp(iconD,L"none")==0)
			return NULL;
		return ::LoadIcon(GetSystemMetrics(SM_CXICON),iconD,modules);
	}
	else if (hIcon)
		return CreateDisabledIcon(hIcon,GetSystemMetrics(SM_CXICON));
	else
		return NULL;
}

// Subclass the renaming editbox to handle Esc and Enter
static LRESULT CALLBACK SubclassEditboxProc( HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData )
{
	if (uMsg==WM_GETDLGCODE)
		return DLGC_WANTALLKEYS;
	return DefSubclassProc(hWnd,uMsg,wParam,lParam);
}

class CSettingsTree: public CWindowImpl<CSettingsTree, CWindow>
{
public:
	BEGIN_MSG_MAP( CSettingsTree )
		MESSAGE_HANDLER( WM_DESTROY, OnDestroy )
		MESSAGE_HANDLER( WM_MOUSEMOVE, OnMouseMove )
		MESSAGE_HANDLER( WM_LBUTTONUP, OnLButtonUp )
		MESSAGE_HANDLER( WM_RBUTTONDOWN, OnRButtonDown )
		MESSAGE_HANDLER( WM_KEYUP, OnKeyUp )
		MESSAGE_HANDLER( WM_GETDLGCODE, OnGetDlgCode )
		MESSAGE_HANDLER( WM_CHAR, OnChar )
		MESSAGE_HANDLER( WM_CAPTURECHANGED, OnCaptureChanged )
		MESSAGE_HANDLER( WM_TIMER, OnTimer )
		REFLECTED_NOTIFY_CODE_HANDLER( TVN_DELETEITEM, OnDeleteItem )
		REFLECTED_NOTIFY_CODE_HANDLER( TVN_BEGINDRAG, OnBeginDrag )
		REFLECTED_NOTIFY_CODE_HANDLER( TVN_BEGINLABELEDIT, OnBeginLabelEdit )
		REFLECTED_NOTIFY_CODE_HANDLER( TVN_ENDLABELEDIT, OnEndLabelEdit )
		REFLECTED_NOTIFY_CODE_HANDLER( TVN_KEYDOWN, OnKeyDown )
	END_MSG_MAP()

	CSettingsTree( void ) { m_ImageList=NULL; }
	void SubclassWindow( HWND hWnd, CCustomTreeDlg *pOwner );
	int LoadIcon( const CTreeItem *pItem );
	int UpdateIcons( CTreeItem *pItem, HTREEITEM hItem );
	void SetDisabled( bool bDisabled ) { m_bDisabled=bDisabled; }
	bool IsDragging( void ) const { return m_DragMode!=DRAG_NONE; }

	void DeleteItem( HTREEITEM hItem );

	void BeginDrag( const CStdCommand *pCommand );
	HTREEITEM CreateStdItem( const CStdCommand *pCommand, HTREEITEM hParent, HTREEITEM hAfter );

	std::vector<HMODULE> m_Modules;

protected:
	LRESULT OnDestroy( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled );
	LRESULT OnMouseMove( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled );
	LRESULT OnLButtonUp( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled );
	LRESULT OnRButtonDown( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled );
	LRESULT OnKeyUp( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled );
	LRESULT OnGetDlgCode( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled );
	LRESULT OnChar( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled );
	LRESULT OnCaptureChanged( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled );
	LRESULT OnTimer( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled );
	LRESULT OnDeleteItem( int idCtrl, LPNMHDR pnmh, BOOL& bHandled );
	LRESULT OnBeginDrag( int idCtrl, LPNMHDR pnmh, BOOL& bHandled );
	LRESULT OnBeginLabelEdit( int idCtrl, LPNMHDR pnmh, BOOL& bHandled );
	LRESULT OnEndLabelEdit( int idCtrl, LPNMHDR pnmh, BOOL& bHandled );
	LRESULT OnKeyDown( int idCtrl, LPNMHDR pnmh, BOOL& bHandled );

private:
	CCustomTreeDlg *m_pOwner;
	std::map<unsigned int,int> m_IconCache;
	HIMAGELIST m_ImageList;
	bool m_bDisabled;

	enum TDragMode
	{
		DRAG_NONE,
		DRAG_DELETE,
		DRAG_ERROR,
		DRAG_MOVE, // these must be the last two
		DRAG_COPY,
	};

	enum TDropLocation
	{
		DROP_NOWHERE,
		DROP_INSIDE,
		DROP_BEFORE,
		DROP_AFTER,
		DROP_LAST,
	};

	enum {
		TIMER_SCROLL=1,
		TIMER_HOVER=2,
	};

	TDragMode m_DragMode;
	HCURSOR m_CursorError, m_CursorMove, m_CursorCopy, m_CursorDel;
	HTREEITEM m_DragItem;
	HTREEITEM m_DropTarget;
	HTREEITEM m_DropSubmenu;
	TDropLocation m_DropLocation;
	int m_Scroll;
	HTREEITEM m_HoverItem;
	const CStdCommand *m_pStdDragItem;
	bool m_bDragLink; // the item being dragged is a link

	bool FindItemByName( HTREEITEM hParent, const wchar_t *name, HTREEITEM hExclude );
	CString CreateUniqueName( CString oldName );
	HTREEITEM CreateDeepCopy( HTREEITEM hSource, HTREEITEM hParent, HTREEITEM hAfter, bool bRename );
};

void CSettingsTree::SubclassWindow( HWND hWnd, 	CCustomTreeDlg *pOwner )
{
	m_pOwner=pOwner;
	int size=GetSystemMetrics(SM_CXSMICON);
	if (!m_ImageList)
		m_ImageList=ImageList_Create(size,size,ILC_COLOR32|ILC_MASK|ILC_MIRROR,1,16);
	CWindowImpl<CSettingsTree, CWindow>::SubclassWindow(hWnd);
	TreeView_SetImageList(hWnd,m_ImageList,TVSIL_NORMAL);

	BITMAPINFO bi={0};
	bi.bmiHeader.biSize=sizeof(BITMAPINFOHEADER);
	bi.bmiHeader.biWidth=bi.bmiHeader.biHeight=size;
	bi.bmiHeader.biPlanes=1;
	bi.bmiHeader.biBitCount=32;
	RECT rc={0,0,size,size};

	HDC hdc=CreateCompatibleDC(NULL);
	HBITMAP bmp=CreateDIBSection(hdc,&bi,DIB_RGB_COLORS,NULL,NULL,0);
	HGDIOBJ bmp0=SelectObject(hdc,bmp);
	FillRect(hdc,&rc,(HBRUSH)(COLOR_WINDOW+1));
	SelectObject(hdc,bmp0);
	DeleteDC(hdc);
	ImageList_AddMasked(m_ImageList,bmp,GetSysColor(COLOR_WINDOW));
	DeleteObject(bmp);

	SHFILEINFO info;
	if (SHGetFileInfo(L"file",FILE_ATTRIBUTE_NORMAL,&info,sizeof(info),SHGFI_USEFILEATTRIBUTES|SHGFI_ICON|SHGFI_SMALLICON))
	{
		ImageList_AddIcon(m_ImageList,info.hIcon);
		DestroyIcon(info.hIcon);
	}
	m_DragMode=DRAG_NONE;
	HMODULE hOle32=LoadLibrary(L"ole32.dll");
	if (hOle32)
	{
		m_CursorError=(HCURSOR)LoadImage(hOle32,MAKEINTRESOURCE(1),IMAGE_CURSOR,0,0,LR_DEFAULTSIZE);
		m_CursorMove=(HCURSOR)LoadImage(hOle32,MAKEINTRESOURCE(2),IMAGE_CURSOR,0,0,LR_DEFAULTSIZE);
		m_CursorCopy=(HCURSOR)LoadImage(hOle32,MAKEINTRESOURCE(3),IMAGE_CURSOR,0,0,LR_DEFAULTSIZE);
		m_CursorDel=(HCURSOR)LoadImage(hOle32,MAKEINTRESOURCE(5),IMAGE_CURSOR,0,0,LR_DEFAULTSIZE);
		FreeLibrary(hOle32);
	}
}

int CSettingsTree::LoadIcon( const CTreeItem *pItem )
{
	unsigned int key=pItem->GetIconKey();
	std::map<unsigned int,int>::const_iterator it=m_IconCache.find(key);
	if (it!=m_IconCache.end())
		return it->second;

	// extract icon
	HICON hIcon=pItem->LoadIcon(true,m_Modules);
	if (hIcon==HICON_NONE)
	{
		m_IconCache[key]=0;
		return 0;
	}
	int index=1;
	if (hIcon)
	{
		index=ImageList_AddIcon(m_ImageList,hIcon);
		DestroyIcon(hIcon);
	}
	m_IconCache[key]=index;
	return index;
}

int CSettingsTree::UpdateIcons( CTreeItem *pItem, HTREEITEM hItem )
{
	TVITEM item={TVIF_IMAGE|TVIF_SELECTEDIMAGE,hItem};

	item.iImage=item.iSelectedImage=LoadIcon(pItem);
	if (hItem)
		TreeView_SetItem(m_hWnd,&item);
	return item.iImage;
}

LRESULT CSettingsTree::OnDestroy( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled )
{
	for (std::vector<HMODULE>::const_iterator it=m_Modules.begin();it!=m_Modules.end();++it)
		FreeLibrary(*it);
	m_Modules.clear();
	if (m_CursorError) DestroyCursor(m_CursorError);
	if (m_CursorMove) DestroyCursor(m_CursorMove);
	if (m_CursorCopy) DestroyCursor(m_CursorCopy);
	if (m_CursorDel) DestroyCursor(m_CursorDel);
	return 0;
}

LRESULT CSettingsTree::OnDeleteItem( int idCtrl, LPNMHDR pnmh, BOOL& bHandled )
{
	// free data
	NMTREEVIEW *pItem=(NMTREEVIEW*)pnmh;
	delete (CTreeItem*)pItem->itemOld.lParam;
	return 0;
}

void CSettingsTree::BeginDrag( const CStdCommand *pCommand )
{
	if (m_bDisabled) return;
	m_pStdDragItem=pCommand;
	m_bDragLink=pCommand->knownFolder!=NULL;
	m_DragItem=TreeView_GetSelection(m_hWnd);
	SetFocus();
	SetCapture();
	m_DragMode=DRAG_COPY;
	m_DropTarget=NULL;
	m_DropLocation=DROP_NOWHERE;
	m_Scroll=0;
	m_HoverItem=NULL;
	m_DropSubmenu=NULL;
	SetCursor(m_CursorCopy);
	TreeView_SelectItem(m_hWnd,NULL);
}

LRESULT CSettingsTree::OnBeginDrag( int idCtrl, LPNMHDR pnmh, BOOL& bHandled )
{
	if (m_bDisabled) return 0;
	NMTREEVIEW *pDrag=(NMTREEVIEW*)pnmh;
	CTreeItem *pItem=(CTreeItem*)pDrag->itemNew.lParam;
	if (!pItem) return 0;
	m_pStdDragItem=NULL;
	m_DragItem=pDrag->itemNew.hItem;
	m_bDragLink=(pItem->pStdCommand && pItem->pStdCommand->knownFolder) || !pItem->link.IsEmpty();
	TreeView_SelectItem(m_hWnd,m_DragItem);
	SetCapture();
	m_DragMode=GetKeyState(VK_CONTROL)<0?DRAG_COPY:DRAG_MOVE;
	m_DropTarget=NULL;
	m_DropLocation=DROP_NOWHERE;
	m_Scroll=0;
	m_HoverItem=NULL;
	m_DropSubmenu=NULL;
	SetCursor(m_DragMode==DRAG_COPY?m_CursorCopy:m_CursorMove);
	SetFocus();
	TreeView_SelectItem(m_hWnd,NULL);
	return 0;
}

LRESULT CSettingsTree::OnMouseMove( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled )
{
	if (m_DragMode==DRAG_NONE)
	{
		bHandled=FALSE;
		return 0;
	}
	// find the drop location
	m_DropTarget=NULL;
	m_DropLocation=DROP_NOWHERE;
	TVHITTESTINFO test={{(short)LOWORD(lParam),(short)HIWORD(lParam)}};
	RECT rc;
	GetClientRect(&rc);
	int h=TreeView_GetItemHeight(m_hWnd);
	int scroll=0;
	if (test.pt.y<h)
		scroll=-1;
	if (test.pt.y>rc.bottom-h)
		scroll=1;
	if (m_Scroll!=scroll)
	{
		m_Scroll=scroll;
		if (m_Scroll)
			SetTimer(TIMER_SCROLL,100);
		else
			KillTimer(TIMER_SCROLL);
	}
	bool bDelete=!PtInRect(&rc,test.pt);
	if (!bDelete)
	{
		HTREEITEM hLast=TreeView_GetLastVisible(m_hWnd);
		if (hLast)
		{
			TreeView_GetItemRect(m_hWnd,hLast,&rc,FALSE);
			if (TreeView_HitTest(m_hWnd,&test))
			{
				m_DropTarget=test.hItem;
			}
			else
			{
				if (test.pt.y>=rc.bottom)
				{
					// after the last visible item
					m_DropTarget=hLast;
				}
			}
		}
		else if (m_pStdDragItem)
			m_DropLocation=DROP_LAST;

		if (m_DropTarget && !m_pStdDragItem)
		{
			// can't drop inside itself
			for (HTREEITEM hItem=TreeView_GetParent(m_hWnd,m_DropTarget);hItem;hItem=TreeView_GetParent(m_hWnd,hItem))
			{
				if (hItem==m_DragItem)
				{
					m_DropTarget=NULL;
					break;
				}
			}
		}
		if (m_DropTarget)
		{
			TreeView_GetItemRect(m_hWnd,m_DropTarget,&rc,FALSE);
			TVITEM item={TVIF_PARAM,m_DropTarget};
			TreeView_GetItem(m_hWnd,&item);
			int dy=test.pt.y-rc.top;
			int h1=(rc.bottom-rc.top)/4;
			int h2=(rc.bottom-rc.top)-h1;
			if (!item.lParam || ((CTreeItem*)item.lParam)->bSeparator)
			{
				h1=(rc.bottom-rc.top)/2;
				h2=(rc.bottom-rc.top)-h1;
			}
			if (dy<h1)
				m_DropLocation=DROP_BEFORE;
			else if (dy>=h2)
				m_DropLocation=DROP_AFTER;
			else
			{
				// can't drop in a separator or itself
				if (!item.lParam || ((CTreeItem*)item.lParam)->bSeparator || (m_DropTarget==m_DragItem && !m_pStdDragItem))
					m_DropTarget=NULL;
				else
					m_DropLocation=DROP_INSIDE;
			}
		}
	}
	// display drop location
	if (!m_DropTarget)
	{
		TreeView_SelectDropTarget(m_hWnd,NULL);
		TreeView_SetInsertMark(m_hWnd,NULL,FALSE);
		if (m_DropLocation==DROP_LAST)
		{
			m_DragMode=DRAG_COPY;
			SetCursor(m_CursorCopy);
		}
		else
		{
			if (m_pStdDragItem)
				bDelete=false;
			m_DragMode=bDelete?DRAG_DELETE:DRAG_ERROR;
			SetCursor(bDelete?m_CursorDel:m_CursorError);
		}
	}
	else
	{
		m_DragMode=(m_pStdDragItem || GetKeyState(VK_CONTROL)<0)?DRAG_COPY:DRAG_MOVE;
		SetCursor(m_DragMode==DRAG_COPY?m_CursorCopy:m_CursorMove);
		if (m_DropLocation==DROP_INSIDE)
		{
			TreeView_SelectDropTarget(m_hWnd,m_DropTarget);
			TreeView_SetInsertMark(m_hWnd,NULL,FALSE);
		}
		else
		{
			TreeView_SelectDropTarget(m_hWnd,NULL);
			TreeView_SetInsertMark(m_hWnd,m_DropTarget,m_DropLocation==DROP_AFTER);
		}
	}

	if (m_DropLocation==DROP_INSIDE)
	{
		if (m_HoverItem!=m_DropTarget)
		{
			SetTimer(TIMER_HOVER,400);
			m_HoverItem=m_DropTarget;
		}
	}
	else if (m_HoverItem)
	{
		m_HoverItem=NULL;
		KillTimer(TIMER_HOVER);
	}
	return 0;
}

bool CSettingsTree::FindItemByName( HTREEITEM hParent, const wchar_t *name, HTREEITEM hExclude )
{
	for (HTREEITEM hItem=hParent?TreeView_GetChild(m_hWnd,hParent):TreeView_GetRoot(m_hWnd);hItem;hItem=TreeView_GetNextSibling(m_hWnd,hItem))
	{
		if (FindItemByName(hItem,name,hExclude))
			return true;
		if (hItem!=hExclude)
		{
			TVITEM item={TVIF_PARAM,hItem};
			TreeView_GetItem(m_hWnd,&item);
			if (item.lParam && _wcsicmp(((CTreeItem*)item.lParam)->name,name)==0)
				return true;
		}
	}
	return false;
}

CString CSettingsTree::CreateUniqueName( CString oldName )
{
	if (!FindItemByName(NULL,oldName,NULL))
		return oldName;
	wchar_t name[256];
	Strcpy(name,_countof(name)-3,oldName);
	// find numeric part
	int pos=Strlen(name);
	while (pos>0 && name[pos-1]>='0' && name[pos-1]<='9')
		pos--;
	for (int i=2;;i++)
	{
		Sprintf(name+pos,_countof(name)-pos,L"%d",i);
		if (!FindItemByName(NULL,name,NULL))
			break;
	}
	return name;
}

HTREEITEM CSettingsTree::CreateDeepCopy( HTREEITEM hSource, HTREEITEM hParent, HTREEITEM hAfter, bool bRename )
{
	TVINSERTSTRUCT insert={hParent,hAfter,{TVIF_PARAM|TVIF_IMAGE|TVIF_SELECTEDIMAGE|TVIF_STATE,hSource,0,TVIS_EXPANDED}};
	TreeView_GetItem(m_hWnd,&insert.item);
	CTreeItem *pSource=(CTreeItem*)insert.item.lParam;
	if (!pSource) return NULL;
	CTreeItem *pNewItem=new CTreeItem(*pSource);
	if (bRename && !pNewItem->bSeparator)
	{
		pNewItem->name=CreateUniqueName(pNewItem->name);
	}
	insert.item.lParam=(LPARAM)pNewItem;
	insert.item.mask|=TVIF_TEXT;
	insert.item.pszText=(LPWSTR)(LPCWSTR)pNewItem->name;
	HTREEITEM hResult=TreeView_InsertItem(m_hWnd,&insert);
	for (HTREEITEM hChild=TreeView_GetChild(m_hWnd,hSource);hChild;hChild=TreeView_GetNextSibling(m_hWnd,hChild))
		CreateDeepCopy(hChild,hResult,TVI_LAST,bRename);
	return hResult;
}

HTREEITEM CSettingsTree::CreateStdItem( const CStdCommand *pCommand, HTREEITEM hParent, HTREEITEM hAfter )
{
	TVINSERTSTRUCT insert={hParent,hAfter,{TVIF_PARAM|TVIF_IMAGE|TVIF_SELECTEDIMAGE|TVIF_STATE|TVIF_TEXT,NULL,TVIS_EXPANDED,TVIS_EXPANDED}};
	CTreeItem *pNewItem=new CTreeItem;
	pNewItem->command=pCommand->name;
	pNewItem->label=pCommand->label;
	pNewItem->tip=pCommand->tip;
	pNewItem->icon=pCommand->icon;
	pNewItem->iconD=pCommand->iconD;
	pNewItem->pStdCommand=pCommand;
	pNewItem->settings=pCommand->settings;

	if (pCommand->itemName)
	{
		pNewItem->bSeparator=false;
		pNewItem->name=pCommand->itemName;
		insert.item.iImage=insert.item.iSelectedImage=UpdateIcons(pNewItem,NULL);
	}
	else
	{
		pNewItem->bSeparator=true;
		pNewItem->name=pCommand->displayName;
	}

	if (!pNewItem->bSeparator)
		pNewItem->name=CreateUniqueName(pNewItem->name);
	insert.item.lParam=(LPARAM)pNewItem;
	insert.item.pszText=(LPWSTR)(LPCWSTR)pNewItem->name;
	return TreeView_InsertItem(m_hWnd,&insert);
}

LRESULT CSettingsTree::OnRButtonDown( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled )
{
	// ignore right-click during dragging
	if (m_DragMode==DRAG_NONE)
		bHandled=FALSE;
	return 0;
}

LRESULT CSettingsTree::OnLButtonUp( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled )
{
	if (m_DragMode!=DRAG_NONE)
	{
		TDragMode mode=m_DragMode;
		HTREEITEM dragItem=m_DragItem;
		if (m_DropSubmenu==m_DropTarget && mode>=DRAG_MOVE)
			m_DropSubmenu=NULL;
		m_DragItem=NULL;
		ReleaseCapture();
		// perform drop operation
		if (mode==DRAG_DELETE)
		{
			TreeView_EnsureVisible(m_hWnd,dragItem);
			if (!m_pStdDragItem)
			{
				TreeView_DeleteItem(m_hWnd,dragItem);
				m_pOwner->SerializeData();
			}
		}
		else if (mode>=DRAG_MOVE)
		{
			HTREEITEM hParent=m_DropTarget;
			HTREEITEM hAfter=TVI_LAST;
			if (m_DropLocation==DROP_AFTER)
			{
				hAfter=hParent;
				hParent=TreeView_GetParent(m_hWnd,hParent);
			}
			else if (m_DropLocation==DROP_BEFORE)
			{
				hAfter=TreeView_GetPrevSibling(m_hWnd,hParent);
				if (!hAfter) hAfter=TVI_FIRST;
				hParent=TreeView_GetParent(m_hWnd,hParent);
			}
			if (hParent)
			{
				HTREEITEM hChild=TreeView_GetChild(m_hWnd,hParent);
				if (hChild)
				{
					TVITEM item={TVIF_PARAM,hChild};
					TreeView_GetItem(m_hWnd,&item);
					if (!item.lParam)
					{
						TreeView_DeleteItem(m_hWnd,hChild);
						hAfter=TVI_LAST;
					}
				}
			}

			HTREEITEM hNewItem;
			if (m_pStdDragItem)
			{
				hNewItem=CreateStdItem(m_pStdDragItem,hParent,hAfter);
			}
			else
			{
				hNewItem=CreateDeepCopy(dragItem,hParent,hAfter,mode==DRAG_COPY);
				if (mode==DRAG_MOVE)
					TreeView_DeleteItem(m_hWnd,dragItem);
			}

			TreeView_SelectItem(m_hWnd,hNewItem);
			TreeView_EnsureVisible(m_hWnd,hNewItem);
			m_pOwner->SerializeData();
		}
	}
	bHandled=FALSE;
	return 0;
}

LRESULT CSettingsTree::OnKeyUp( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled )
{
	if (m_DragMode>=DRAG_MOVE)
	{
		m_DragMode=(m_pStdDragItem || GetKeyState(VK_CONTROL)<0)?DRAG_COPY:DRAG_MOVE;
		SetCursor(m_DragMode==DRAG_COPY?m_CursorCopy:m_CursorMove);
	}
	bHandled=FALSE;
	return 0;
}

LRESULT CSettingsTree::OnGetDlgCode( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled )
{
	if (m_DragMode!=DRAG_NONE || (lParam && ((MSG*)lParam)->message==WM_KEYDOWN && ((MSG*)lParam)->wParam==VK_RETURN))
		return DLGC_WANTALLKEYS;
	bHandled=FALSE;
	return 0;
}

LRESULT CSettingsTree::OnChar( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled )
{
	if (wParam==VK_RETURN)
		return 0;
	bHandled=FALSE;
	return 0;
}

LRESULT CSettingsTree::OnCaptureChanged( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled )
{
	if (m_DragMode!=DRAG_NONE)
	{
		m_DragMode=DRAG_NONE;
		TreeView_SelectDropTarget(m_hWnd,NULL);
		TreeView_SetInsertMark(m_hWnd,NULL,FALSE);
		SetCursor(LoadCursor(NULL,IDC_ARROW));
		if (m_DragItem) TreeView_SelectItem(m_hWnd,m_DragItem);
		KillTimer(TIMER_SCROLL);
		KillTimer(TIMER_HOVER);
		if (m_DropSubmenu)
			TreeView_DeleteItem(m_hWnd,m_DropSubmenu);
		m_DropSubmenu=NULL;
		m_HoverItem=NULL;
	}
	bHandled=FALSE;
	return 0;
}

LRESULT CSettingsTree::OnTimer( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled )
{
	if (wParam==TIMER_SCROLL)
	{
		if (m_Scroll>0)
			SendMessage(WM_VSCROLL,SB_LINEDOWN);
		else
			SendMessage(WM_VSCROLL,SB_LINEUP);
	}
	else if (wParam==TIMER_HOVER)
	{
		if (m_HoverItem)
		{
			TreeView_Expand(m_hWnd,m_HoverItem,TVE_EXPAND);
			if (!TreeView_GetChild(m_hWnd,m_HoverItem))
			{
				if (m_DropSubmenu)
					TreeView_DeleteItem(m_hWnd,m_DropSubmenu);
				TVINSERTSTRUCT insert={m_HoverItem,TVI_FIRST,{TVIF_PARAM|TVIF_IMAGE|TVIF_SELECTEDIMAGE|TVIF_STATE|TVIF_TEXT,NULL,TVIS_EXPANDED,TVIS_EXPANDED}};
				CString text=LoadStringEx(IDS_EMPTY_MENU);
				insert.item.pszText=(LPWSTR)(LPCWSTR)text;
				m_DropSubmenu=TreeView_InsertItem(m_hWnd,&insert);
				TreeView_SelectItem(m_hWnd,m_DropSubmenu);
				TreeView_EnsureVisible(m_hWnd,m_DropSubmenu);
			}
		}
		KillTimer(TIMER_HOVER);
	}
	else
		bHandled=FALSE;
	return 0;
}

LRESULT CSettingsTree::OnBeginLabelEdit( int idCtrl, LPNMHDR pnmh, BOOL& bHandled )
{
	if (m_bDisabled) return TRUE;
	SetWindowSubclass(TreeView_GetEditControl(m_hWnd),SubclassEditboxProc,'CLSH',0);
	NMTVDISPINFO *pInfo=(NMTVDISPINFO*)pnmh;
	if (!pInfo->item.lParam || ((CTreeItem*)pInfo->item.lParam)->bSeparator)
		return TRUE;
	return FALSE;
}

LRESULT CSettingsTree::OnEndLabelEdit( int idCtrl, LPNMHDR pnmh, BOOL& bHandled )
{
	NMTVDISPINFO *pInfo=(NMTVDISPINFO*)pnmh;
	if (pInfo->item.pszText)
	{
		if (!*pInfo->item.pszText)
		{
			::MessageBox(m_hWnd,LoadStringEx(IDS_ERROR_EMPTY),LoadStringEx(IDS_ERROR_TITLE),MB_OK|MB_ICONERROR);
			return FALSE;
		}
		bool bValid=true;
		for (const wchar_t *c=pInfo->item.pszText;*c;c++)
		{
			if ((*c>='a' && *c<='z') || (*c>='A' && *c<='Z') || (*c>='0' && *c<='9') || *c=='_')
				continue;
			::MessageBox(m_hWnd,LoadStringEx(IDS_ERROR_ASCII),LoadStringEx(IDS_ERROR_TITLE),MB_OK|MB_ICONERROR);
			return FALSE;
		}
		wchar_t text[1024];
		if (m_pOwner->IsSeparator(pInfo->item.pszText))
		{
			Sprintf(text,_countof(text),LoadStringEx(IDS_ERROR_SEPARATOR),pInfo->item.pszText);
			::MessageBox(m_hWnd,text,LoadStringEx(IDS_ERROR_TITLE),MB_OK|MB_ICONERROR);
			return FALSE;
		}
		if (FindItemByName(NULL,pInfo->item.pszText,pInfo->item.hItem))
		{
			Sprintf(text,_countof(text),LoadStringEx(IDS_DUPLICATE_ITEM),pInfo->item.pszText);
			::MessageBox(m_hWnd,text,LoadStringEx(IDS_ERROR_TITLE),MB_OK|MB_ICONERROR);
			return FALSE;
		}
		((CTreeItem*)pInfo->item.lParam)->name=pInfo->item.pszText;
		m_pOwner->SerializeData();
		return TRUE;
	}
	return FALSE;
}

LRESULT CSettingsTree::OnKeyDown( int idCtrl, LPNMHDR pnmh, BOOL& bHandled )
{
	NMTVKEYDOWN *pKey=(NMTVKEYDOWN*)pnmh;
	if (m_DragMode!=DRAG_NONE)
	{
		if (pKey->wVKey==VK_ESCAPE)
		{
			ReleaseCapture();
			return TRUE;
		}
		if (m_DragMode>=DRAG_MOVE)
		{
			m_DragMode=(m_pStdDragItem || GetKeyState(VK_CONTROL)<0)?DRAG_COPY:DRAG_MOVE;
			SetCursor(m_DragMode==DRAG_COPY?m_CursorCopy:m_CursorMove);
		}
	}
	// F2 to rename
	if (pKey->wVKey==VK_F2)
	{
		HTREEITEM hSelected=TreeView_GetSelection(m_hWnd);
		if (hSelected)
			TreeView_EditLabel(m_hWnd,hSelected);
	}
	// Del to delete
	if (pKey->wVKey==VK_DELETE)
	{
		HTREEITEM hSelected=TreeView_GetSelection(m_hWnd);
		if (hSelected)
			DeleteItem(hSelected);
	}
	return 0;
}

void CSettingsTree::DeleteItem( HTREEITEM hItem )
{
	if (!m_bDisabled)
	{
		TreeView_DeleteItem(m_hWnd,hItem);
		m_pOwner->SerializeData();
	}
}

///////////////////////////////////////////////////////////////////////////////

class CCommandsTree: public CWindowImpl<CCommandsTree, CWindow>
{
public:
	BEGIN_MSG_MAP( CCommandsTree )
		MESSAGE_HANDLER( WM_GETDLGCODE, OnGetDlgCode )
		MESSAGE_HANDLER( WM_CHAR, OnChar )
	END_MSG_MAP()

protected:
	LRESULT OnGetDlgCode( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled );
	LRESULT OnChar( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled );
};

LRESULT CCommandsTree::OnGetDlgCode( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled )
{
	if (lParam && ((MSG*)lParam)->message==WM_KEYDOWN && ((MSG*)lParam)->wParam==VK_RETURN)
		return DLGC_WANTALLKEYS;
	bHandled=FALSE;
	return 0;
}

LRESULT CCommandsTree::OnChar( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled )
{
	if (wParam==VK_RETURN)
		return 0;
	bHandled=FALSE;
	return 0;
}

///////////////////////////////////////////////////////////////////////////////

CEditCustomItemDlg::~CEditCustomItemDlg( void )
{
	if (m_hIcon) DestroyIcon(m_hIcon);
	if (m_hIconD) DestroyIcon(m_hIconD);
}

void CEditCustomItemDlg::InitDialog( CWindow commandCombo, const CStdCommand *pStdcommands, CWindow linkCombo, const KNOWNFOLDERID *const *pCommonLinks )
{
	CResizeableDlg<CEditCustomItemDlg>::InitResize(MOVE_HORIZONTAL|MOVE_GRIPPER);
	CString str;
	GetWindowText(str);
	CString title;
	title.Format(str,m_pItem->name);
	SetWindowText(title);

	for (int i=0;pStdcommands[i].name;i++)
		if (pStdcommands[i].itemName && *pStdcommands[i].name && *pStdcommands[i].displayName)
			commandCombo.SendMessage(CB_ADDSTRING,0,(LPARAM)pStdcommands[i].name);

	commandCombo.SetWindowText(m_pItem->command);
	int idx=(int)commandCombo.SendMessage(CB_FINDSTRINGEXACT,-1,(LPARAM)(const wchar_t*)m_pItem->command);
	if (idx>=0)
		commandCombo.SendMessage(CB_SETCURSEL,idx);

	for (int i=0;pCommonLinks[i];i++)
	{
		if (*pCommonLinks[i]==FOLDERID_DesktopRoot)
		{
			linkCombo.SendMessage(CB_ADDSTRING,0,(LPARAM)L"Main Desktop | ::{Desktop}");
		}
		else
		{
			PIDLIST_ABSOLUTE path;
			if (FAILED(SHGetKnownFolderIDList(*pCommonLinks[i],0,NULL,&path)) || !path) continue;
			wchar_t *pName;
			if (SUCCEEDED(SHGetNameFromIDList(path,SIGDN_PARENTRELATIVEEDITING,&pName)))
			{
				wchar_t *pPath;
				if (SUCCEEDED(SHGetNameFromIDList(path,SIGDN_DESKTOPABSOLUTEPARSING,&pPath)))
				{
					wchar_t text[_MAX_PATH*2];
					int len=Sprintf(text,_countof(text),L"%s | ",pName);
					if (!PathUnExpandEnvStrings(pPath,text+len,_countof(text)-len))
						Strcpy(text+len,_countof(text)-len,pPath);
					linkCombo.SendMessage(CB_ADDSTRING,0,(LPARAM)text);
					CoTaskMemFree(pPath);
				}
				CoTaskMemFree(pName);
			}
			ILFree(path);
		}
	}

	linkCombo.SetWindowText(m_pItem->link);
	idx=(int)linkCombo.SendMessage(CB_FINDSTRINGEXACT,-1,(LPARAM)(const wchar_t*)m_pItem->link);
	if (idx>=0)
		linkCombo.SendMessage(CB_SETCURSEL,idx);

	m_hIcon=NULL;
	m_IconKey=0;
	m_hIconD=NULL;
	m_IconDKey=0;
	m_StoredItem=*m_pItem;

	RECT rc;
	CRegKey regSettings;
	if (regSettings.Open(HKEY_CURRENT_USER,GetSettingsRegPath())==ERROR_SUCCESS)
	{
		ULONG size=sizeof(rc);
		if (regSettings.QueryBinaryValue(L"CEditCustomItemDlg",&rc,&size)==ERROR_SUCCESS && size==sizeof(rc))
			SetStoreRect(rc);
	}

	COMBOBOXINFO info={sizeof(info)};
	commandCombo.SendMessage(CB_GETCOMBOBOXINFO,0,(LPARAM)&info);
	SHAutoComplete(info.hwndItem,SHACF_FILESYSTEM);
	linkCombo.SendMessage(CB_GETCOMBOBOXINFO,0,(LPARAM)&info);
	SHAutoComplete(info.hwndItem,SHACF_FILESYSTEM);
}

LRESULT CEditCustomItemDlg::OnSize( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled )
{
	CResizeableDlg<CEditCustomItemDlg>::OnSize();
	return 0;
}

void CEditCustomItemDlg::StorePlacement( void )
{
	RECT rc;
	GetStoreRect(rc);

	CRegKey regSettings;
	if (regSettings.Open(HKEY_CURRENT_USER,GetSettingsRegPath())!=ERROR_SUCCESS)
		regSettings.Create(HKEY_CURRENT_USER,GetSettingsRegPath());
	regSettings.SetBinaryValue(L"CEditCustomItemDlg",&rc,sizeof(rc));
}

LRESULT CEditCustomItemDlg::OnOK( WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled )
{
	::EnableWindow(m_EnableParent,TRUE);
	m_bResult=true;
	StorePlacement();
	DestroyWindow();
	return 0;
}

LRESULT CEditCustomItemDlg::OnCancel( WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled )
{
	::EnableWindow(m_EnableParent,TRUE);
	m_bResult=false;
	StorePlacement();
	DestroyWindow();
	(*m_pItem)=m_StoredItem;
	return 0;
}

void CEditCustomItemDlg::UpdateIcons( int iconID, int iconDID )
{
	unsigned int key=m_pItem->GetIconKey();

	if (key!=m_IconKey)
	{
		m_IconKey=key;
		if (m_hIcon) DestroyIcon(m_hIcon);
		m_hIcon=m_pItem->LoadIcon(false,m_Modules);
		SendDlgItemMessage(iconID,STM_SETICON,(WPARAM)m_hIcon);
	}

	if (iconDID)
	{
		key=m_pItem->GetIconDKey(key);
		if (key!=m_IconDKey)
		{
			m_IconDKey=key;
			if (m_hIconD) DestroyIcon(m_hIconD);
			m_hIconD=m_pItem->LoadIconD(m_hIcon,m_Modules);
			SendDlgItemMessage(iconDID,STM_SETICON,(WPARAM)m_hIconD);
		}
	}
}

CString CEditCustomItemDlg::GetComboText( WORD wNotifyCode, WORD wID )
{
	CString text;
	if (wNotifyCode==CBN_SELENDOK)
	{
		int idx=(int)SendDlgItemMessage(wID,CB_GETCURSEL);
		text.Empty();
		if (idx>=0)
		{
			int len=(int)SendDlgItemMessage(wID,CB_GETLBTEXTLEN,idx);
			SendDlgItemMessage(wID,CB_GETLBTEXT,idx,(LPARAM)text.GetBuffer(len+1));
			text.ReleaseBuffer(len);
		}
	}
	else
		GetDlgItemText(wID,text);
	text.TrimLeft();
	text.TrimRight();
	return text;
}

///////////////////////////////////////////////////////////////////////////////

class CBrowseLinkEvents: public IFileDialogEvents, public IFileDialogControlEvents
{
public:
	CBrowseLinkEvents( void ) { m_pResult=NULL; }

	// IUnknown
	virtual HRESULT STDMETHODCALLTYPE QueryInterface( REFIID riid, void **ppvObject );
	virtual ULONG STDMETHODCALLTYPE AddRef( void ) { return 1; }
	virtual ULONG STDMETHODCALLTYPE Release( void ) { return 1; }

	// IFileDialogEvents
	STDMETHOD(OnFileOk)( IFileDialog *pfd ) { return S_OK; }
	STDMETHOD(OnFolderChanging)( IFileDialog *pfd, IShellItem *psiFolder ) { return S_OK; }
	STDMETHOD(OnFolderChange)( IFileDialog *pfd ) { return S_OK; }
	STDMETHOD(OnSelectionChange)( IFileDialog *pfd ) { return S_OK; }
	STDMETHOD(OnShareViolation)( IFileDialog *pfd, IShellItem *psi, FDE_SHAREVIOLATION_RESPONSE *pResponse ) { return S_OK; }
	STDMETHOD(OnTypeChange)( IFileDialog *pfd ) { return S_OK; }
	STDMETHOD(OnOverwrite)( IFileDialog *pfd, IShellItem *psi, FDE_OVERWRITE_RESPONSE *pResponse ) { return S_OK; }

	// IFileDialogControlEvents
	STDMETHOD(OnItemSelected)( IFileDialogCustomize *pfdc, DWORD dwIDCtl, DWORD dwIDItem ) { return S_OK; }
	STDMETHOD(OnButtonClicked)( IFileDialogCustomize *pfdc, DWORD dwIDCtl );
	STDMETHOD(OnCheckButtonToggled)( IFileDialogCustomize *pfdc, DWORD dwIDCtl, BOOL bChecked ) { return S_OK; }
	STDMETHOD(OnControlActivating)( IFileDialogCustomize *pfdc, DWORD dwIDCtl ) { return S_OK; }

	CComPtr<IShellItem> m_pResult;
};

HRESULT STDMETHODCALLTYPE CBrowseLinkEvents::QueryInterface( REFIID riid, void **ppvObject )
{
	if (riid==IID_IUnknown)
	{
		*ppvObject=(IUnknown*)(IFileDialogEvents*)this;
		return S_OK;
	}
	if (riid==IID_IFileDialogEvents)
	{
		*ppvObject=(IFileDialogEvents*)this;
		return S_OK;
	}
	if (riid==IID_IFileDialogControlEvents)
	{
		*ppvObject=(IFileDialogControlEvents*)this;
		return S_OK;
	}
	*ppvObject=NULL;
	return E_NOINTERFACE;
}

HRESULT STDMETHODCALLTYPE CBrowseLinkEvents::OnButtonClicked( IFileDialogCustomize *pfdc, DWORD dwIDCtl )
{
	CComQIPtr<IFileDialog> pfd=pfdc;
	if (!pfd) return E_INVALIDARG;

	CComPtr<IShellItem> pItem;
	if (SUCCEEDED(pfd->GetCurrentSelection(&pItem)))
	{
		SFGAOF attr;
		if (SUCCEEDED(pItem->GetAttributes(SFGAO_FOLDER|SFGAO_STREAM,&attr)) && (attr&(SFGAO_FOLDER|SFGAO_STREAM))!=SFGAO_FOLDER)
			pItem=NULL;
	}
	if (!pItem)
	{
		pfd->GetFolder(&pItem);
	}
	m_pResult=pItem;
	pfd->Close(S_FALSE);
	return S_OK;
}

///////////////////////////////////////////////////////////////////////////////

bool CEditCustomItemDlg::BrowseCommand( HWND parent, wchar_t *text )
{
	if (text[0]=='"')
	{
		// remove quotes
		int len=Strlen(text);
		if (len>1 && text[len-1]=='"')
		{
			len-=2;
			memmove(text,text+1,len*2);
			text[len]=0;
		}
	}
	OPENFILENAME ofn={sizeof(ofn)};
	ofn.hwndOwner=m_hWnd;
	wchar_t filters[256];
	Strcpy(filters,_countof(filters),LoadStringEx(IDS_COMMAND_FILTERS));
	for (wchar_t *c=filters;*c;c++)
		if (*c=='|') *c=0;
	ofn.lpstrFilter=filters;
	ofn.nFilterIndex=1;
	DoEnvironmentSubst(text,_MAX_PATH);
	ofn.lpstrFile=text;
	ofn.nMaxFile=_MAX_PATH;
	ofn.Flags=OFN_DONTADDTORECENT|OFN_ENABLESIZING|OFN_EXPLORER|OFN_FILEMUSTEXIST|OFN_HIDEREADONLY|OFN_NOCHANGEDIR;
	bool bRes=false;
	if (GetOpenFileName(&ofn))
	{
		bRes=true;
	}
	else if (CommDlgExtendedError()==FNERR_INVALIDFILENAME)
	{
		// bad path. clear path and try again
		text[0]=0;
		if (GetOpenFileName(&ofn))
			bRes=true;
	}
	if (bRes)
	{
		wchar_t buf[_MAX_PATH];
		UnExpandEnvStrings(text,buf,_countof(buf));
		// add quotes if needed
		if (wcschr(buf,' '))
		{
			int len=Strlen(buf);
			if (len<_MAX_PATH-3)
			{
				memcpy(text+1,buf,len*2);
				text[0]='"';
				text[len+1]='"';
				text[len+2]=0;
			}
		}
		else
			Strcpy(text,_MAX_PATH,buf);
	}
	return bRes;
}

bool CEditCustomItemDlg::BrowseLink( HWND parent, wchar_t *text )
{
	DoEnvironmentSubst(text,_MAX_PATH);

	CComPtr<IFileOpenDialog> pDialog;
	if (FAILED(pDialog.CoCreateInstance(CLSID_FileOpenDialog,NULL,CLSCTX_INPROC_SERVER)))
		return false;
	CComQIPtr<IFileDialogCustomize> pCustomize=pDialog;
	if (!pCustomize)
		return false;

	pDialog->SetTitle(LoadStringEx(IDS_PICK_LINK_TITLE));
	pDialog->SetOkButtonLabel(LoadStringEx(IDS_PICK_LINK_FILE));
	wchar_t button[256];
	Sprintf(button,_countof(button),L"  %s  ",LoadStringEx(IDS_PICK_LINK_FOLDER));
	pCustomize->AddPushButton(101,button);

	CBrowseLinkEvents events;
	DWORD cookie;
	pDialog->Advise(&events,&cookie);
	pDialog->SetOptions(FOS_ALLNONSTORAGEITEMS|FOS_FILEMUSTEXIST|FOS_DONTADDTORECENT|FOS_DEFAULTNOMINIMODE);
	{
		const wchar_t *c=wcschr(text,'|');
		if (c)
		{
			for (c++;*c==' ';)
				c++;
		}
		else
			c=text;
		CComPtr<IShellItem> pItem;
		if (SUCCEEDED(SHCreateItemFromParsingName(c,NULL,IID_IShellItem,(void**)&pItem)))
			pDialog->SetFolder(pItem);
	}

	HRESULT hr=pDialog->Show(parent);
	CComPtr<IShellItem> pResult;
	if (hr==S_OK)
		pDialog->GetResult(&pResult);
	else if (hr==S_FALSE)
		pResult=events.m_pResult;

	pDialog->Unadvise(cookie);

	if (pResult)
	{
		wchar_t *pName;
		pResult->GetDisplayName(SIGDN_DESKTOPABSOLUTEPARSING,&pName);
		UnExpandEnvStrings(pName,text,_MAX_PATH);
		CoTaskMemFree(pName);
	}

	return pResult!=NULL;
}

bool CEditCustomItemDlg::BrowseIcon( wchar_t *text )
{
	int id=0;
	if (*text)
	{
		wchar_t *c=wcsrchr(text,',');
		if (c)
		{
			*c=0;
			id=_wtol(c+1);
		}
	}
	else
		Strcpy(text,_MAX_PATH,L"shell32.dll");
	if (BrowseForIcon(m_hWnd,text,id))
	{
		if (_wcsicmp(PathFindExtension(text),L".ico")==0)
			return true;
		if (id>0)
		{
			wchar_t buf[20];
			Sprintf(buf,_countof(buf),L", %d",id);
			Strcat(text,_MAX_PATH,buf);
			return true;
		}
	}
	return false;
}

bool CEditCustomItemDlg::Run( HWND parent, int dlgID )
{
	parent=GetAncestor(parent,GA_ROOT);
	DLGTEMPLATE *pTemplate=LoadDialogEx(dlgID);
	Create(parent,pTemplate);
	ShowWindow(SW_SHOWNORMAL);
	::EnableWindow(parent,FALSE);
	SetEnableParent(parent);
	MSG msg;
	while (m_hWnd && GetMessage(&msg,0,0,0))
	{
		if (IsDialogMessage(&msg)) continue;
		TranslateMessage(&msg);
		DispatchMessage(&msg);
	}
	return GetResult();
}

///////////////////////////////////////////////////////////////////////////////

// Subclass the tooltip to delay the tip when the mouse moves from one tree item to the next
static LRESULT CALLBACK SubclassInfoTipProc( HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData )
{
	if (uMsg==TTM_UPDATE)
	{
		int time=(int)SendMessage(hWnd,TTM_GETDELAYTIME,TTDT_RESHOW,0);
		SetTimer(hWnd,'CLSH',time,NULL);
		return 0;
	}
	if (uMsg==WM_TIMER && wParam=='CLSH')
	{
		KillTimer(hWnd,wParam);
		DefSubclassProc(hWnd,TTM_UPDATE,0,0);
		return 0;
	}
	return DefSubclassProc(hWnd,uMsg,wParam,lParam);
}

LRESULT CCustomTreeDlg::OnInitDialog( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled )
{
	m_pSetting=NULL;

	CResizeableDlg<CCustomTreeDlg>::InitResize();
	EnableThemeDialogTexture(m_hWnd,ETDT_ENABLETAB);
	m_Tree.SubclassWindow(GetDlgItem(IDC_TREEITEMS),this);
	m_CommandsTree.SubclassWindow(GetDlgItem(IDC_TREECOMMANDS));
	TreeView_SetImageList(m_CommandsTree,TreeView_GetImageList(m_Tree,TVSIL_NORMAL),TVSIL_NORMAL);
	if (IsAppThemed())
	{
		m_Tree.SetWindowLong(GWL_STYLE,m_Tree.GetWindowLong(GWL_STYLE)|TVS_TRACKSELECT);
		SetWindowTheme(m_Tree,L"Explorer",NULL);
		m_CommandsTree.SetWindowLong(GWL_STYLE,m_CommandsTree.GetWindowLong(GWL_STYLE)|TVS_TRACKSELECT);
		SetWindowTheme(m_CommandsTree,L"Explorer",NULL);
	}

	int height=TreeView_GetItemHeight(m_Tree);
	int minHeight=GetSystemMetrics(SM_CXSMICON)+2;
	if (height<minHeight)
	{
		TreeView_SetItemHeight(m_Tree,minHeight);
		TreeView_SetItemHeight(m_CommandsTree,minHeight);
	}

	HWND tooltip=TreeView_GetToolTips(m_CommandsTree);
	SendMessage(tooltip,TTM_SETDELAYTIME,TTDT_AUTOPOP,10000);
	SendMessage(tooltip,TTM_SETDELAYTIME,TTDT_INITIAL,1000);
	SendMessage(tooltip,TTM_SETDELAYTIME,TTDT_RESHOW,1000);
	SetWindowSubclass(tooltip,SubclassInfoTipProc,'CLSH',0);

	tooltip=TreeView_GetToolTips(m_Tree);
	SendMessage(tooltip,TTM_SETDELAYTIME,TTDT_AUTOPOP,10000);
	SendMessage(tooltip,TTM_SETDELAYTIME,TTDT_INITIAL,1000);
	SendMessage(tooltip,TTM_SETDELAYTIME,TTDT_RESHOW,1000);
	SetWindowSubclass(tooltip,SubclassInfoTipProc,'CLSH',0);

	for (int i=0;m_pStdCommands[i].name;i++)
	{
		const CStdCommand &command=m_pStdCommands[i];
		if (!*command.displayName) continue;
		TVINSERTSTRUCT insert={NULL,TVI_LAST,{TVIF_TEXT|TVIF_PARAM|TVIF_IMAGE|TVIF_SELECTEDIMAGE,NULL,0,0,(LPWSTR)(LPCWSTR)command.displayName}};
		insert.item.iImage=insert.item.iSelectedImage=1;
		if (!command.icon)
			insert.item.iImage=insert.item.iSelectedImage=0;
		else
		{
			CTreeItem tItem;
			tItem.icon=command.icon;
			tItem.pStdCommand=&command;
			insert.item.iImage=insert.item.iSelectedImage=m_Tree.LoadIcon(&tItem);
		}

		insert.item.lParam=(LPARAM)&command;
		TreeView_InsertItem(m_CommandsTree,&insert);
	}

	CString str=LoadStringEx(m_bMenu?IDS_TREE_MENU:IDS_TREE_TOOLBAR);
	SetDlgItemText(IDC_STATICLEFT,str);
	str=LoadStringEx(m_bMenu?IDS_DRAG_DROP_SM_TIP:IDS_DRAG_DROP_EXP_TIP);
	SetDlgItemText(IDC_STATICHINT,str);

	return TRUE;
}

LRESULT CCustomTreeDlg::OnSize( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled )
{
	CResizeableDlg<CCustomTreeDlg>::OnSize();
	return 0;
}

LRESULT CCustomTreeDlg::OnContextMenu( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled )
{
	CWindow tree=(HWND)wParam;
	if (tree.m_hWnd!=m_Tree.m_hWnd && tree.m_hWnd!=m_CommandsTree.m_hWnd)
		return 0;
	if (m_Tree.IsDragging())
	{
		ReleaseCapture();
		return 0;
	}
	HTREEITEM hItem=NULL;
	POINT pt;
	if ((DWORD)lParam==0xFFFFFFFF)
	{
		// keyboard
		hItem=TreeView_GetSelection(tree);
		if (hItem)
		{
			RECT rc;
			TreeView_GetItemRect(tree,hItem,&rc,TRUE);
			pt.x=(rc.left+rc.right)/2;
			pt.y=(rc.top+rc.bottom)/2;
		}
		else
			pt.x=pt.y=0;
		tree.ClientToScreen(&pt);
	}
	else
	{
		DWORD pos=GetMessagePos();
		pt.x=(short)LOWORD(pos);
		pt.y=(short)HIWORD(pos);
		TVHITTESTINFO test={pt};
		tree.ScreenToClient(&test.pt);
		if (TreeView_HitTest(tree,&test) && (test.flags&(TVHT_ONITEMICON|TVHT_ONITEMLABEL)))
		{
			TreeView_SelectItem(tree,test.hItem);
			hItem=test.hItem;
		}
	}

	CTreeItem *pItem=NULL;
	if (tree.m_hWnd==m_Tree.m_hWnd)
	{
		if (hItem)
		{
			TVITEM item={TVIF_PARAM,hItem};
			TreeView_GetItem(m_Tree,&item);
			pItem=(CTreeItem*)item.lParam;
		}
	}
	else if (!hItem)
		return 0;

	enum {
		CMD_EDIT=1,
		CMD_RENAME,
		CMD_DELETE,
		CMD_RESETALL,
		CMD_ADD,
		CMD_ADD_MENU,
	};

	HMENU menu=CreatePopupMenu();
	wchar_t text[256];
	if (tree.m_hWnd==m_Tree.m_hWnd)
	{
		if (pItem)
		{
			if (!pItem->bSeparator)
			{
				Sprintf(text,_countof(text),L"%s\tEnter",LoadStringEx(IDS_MENU_EDIT));
				AppendMenu(menu,MF_STRING,CMD_EDIT,text);
				SetMenuDefaultItem(menu,CMD_EDIT,FALSE);
				Sprintf(text,_countof(text),L"%s\tF2",LoadStringEx(IDS_MENU_RENAME));
				AppendMenu(menu,MF_STRING,CMD_RENAME,text);
				if (m_pSetting->flags&CSetting::FLAG_LOCKED_MASK)
				{
					EnableMenuItem(menu,CMD_EDIT,MF_BYCOMMAND|MF_GRAYED);
					EnableMenuItem(menu,CMD_RENAME,MF_BYCOMMAND|MF_GRAYED);
				}
				if (!TreeView_GetChild(m_Tree,hItem))
					AppendMenu(menu,MF_STRING,CMD_ADD_MENU,LoadStringEx(IDS_INSERT_MENU));
				if (m_pSetting->flags&CSetting::FLAG_LOCKED_MASK)
					EnableMenuItem(menu,CMD_ADD_MENU,MF_BYCOMMAND|MF_GRAYED);
			}
		}
		Sprintf(text,_countof(text),L"%s\tDel",LoadStringEx(IDS_MENU_DELETE));
		AppendMenu(menu,MF_STRING,CMD_DELETE,text);
		if (m_pSetting->flags&CSetting::FLAG_LOCKED_MASK)
			EnableMenuItem(menu,CMD_DELETE,MF_BYCOMMAND|MF_GRAYED);
		if (!(m_pSetting->flags&CSetting::FLAG_LOCKED_MASK))
			AppendMenu(menu,MF_STRING,CMD_RESETALL,LoadStringEx(m_bMenu?IDS_RESET_MENU:IDS_RESET_TOOLBAR));
	}
	else
	{
		Sprintf(text,_countof(text),L"%s\tEnter",LoadStringEx(m_bMenu?IDS_ADD_MENU:IDS_ADD_TOOLBAR));
		AppendMenu(menu,MF_STRING,CMD_ADD,text);
		SetMenuDefaultItem(menu,CMD_ADD,FALSE);
		if (m_pSetting->flags&CSetting::FLAG_LOCKED_MASK)
			EnableMenuItem(menu,CMD_ADD,MF_BYCOMMAND|MF_GRAYED);
	}

	int res=TrackPopupMenu(menu,TPM_RETURNCMD,pt.x,pt.y,0,m_hWnd,NULL);
	DestroyMenu(menu);
	if (res==CMD_EDIT)
		EditItemInternal(pItem,hItem);
	if (res==CMD_DELETE)
		m_Tree.DeleteItem(hItem);
	if (res==CMD_RENAME)
		TreeView_EditLabel(m_Tree,hItem);
	if (res==CMD_RESETALL)
	{
		if (::MessageBox(m_hWnd,LoadStringEx(m_bMenu?IDS_RESET_MENU_WARN:IDS_RESET_TOOLBAR_WARN),LoadStringEx(m_bMenu?IDS_RESET_MENU:IDS_RESET_TOOLBAR),MB_YESNO|MB_ICONWARNING)==IDYES)
		{
			{
				CSettingsLockWrite lock;
				m_pSetting->value=m_pSetting->defValue;
				m_pSetting->flags|=CSetting::FLAG_DEFAULT;
			}
			SetGroup(m_pSetting-1,true);
		}
	}
	if (res==CMD_ADD)
	{
		AddItem(hItem);
	}
	if (res==CMD_ADD_MENU)
	{
		TVINSERTSTRUCT insert={hItem,TVI_FIRST,{TVIF_PARAM|TVIF_IMAGE|TVIF_SELECTEDIMAGE|TVIF_STATE|TVIF_TEXT,NULL,TVIS_EXPANDED,TVIS_EXPANDED}};
		CString text=LoadStringEx(IDS_EMPTY_MENU);
		insert.item.pszText=(LPWSTR)(LPCWSTR)text;
		HTREEITEM hNew=TreeView_InsertItem(m_Tree,&insert);
		TreeView_SelectItem(m_Tree,hNew);
		TreeView_EnsureVisible(m_Tree,hNew);
	}
	return 0;
}

bool CCustomTreeDlg::IsSeparator( const wchar_t *name )
{
	for (int i=0;m_pStdCommands[i].name;i++)
		if (_wcsicmp(m_pStdCommands[i].name,name)==0)
			return !m_pStdCommands[i].itemName;
	return false;
}

LRESULT CCustomTreeDlg::OnGetInfoTip( int idCtrl, LPNMHDR pnmh, BOOL& bHandled )
{
	NMTVGETINFOTIP *pTip=(NMTVGETINFOTIP*)pnmh;
	const CStdCommand *pCommand=NULL;
	if (idCtrl==IDC_TREECOMMANDS)
		pCommand=(CStdCommand*)pTip->lParam;
	else if (!(m_pSetting->flags&CSetting::FLAG_LOCKED_REG))
	{
		CTreeItem *pItem=(CTreeItem*)pTip->lParam;
		if (pItem)
			pCommand=pItem->pStdCommand;
	}
	if (pCommand)
	{
		if (idCtrl==IDC_TREECOMMANDS)
			Sprintf(pTip->pszText,pTip->cchTextMax,L"%s\n%s",LoadStringEx(pCommand->tipID),LoadStringEx(m_bMenu?IDS_COMMAND_SM_TIP:IDS_COMMAND_EXP_TIP));
		else
			Strcpy(pTip->pszText,pTip->cchTextMax,LoadStringEx(pCommand->tipID));
	}
	return 0;
}

LRESULT CCustomTreeDlg::OnBeginDrag( int idCtrl, LPNMHDR pnmh, BOOL& bHandled )
{
	NMTREEVIEW *pDrag=(NMTREEVIEW*)pnmh;
	m_Tree.BeginDrag((CStdCommand*)pDrag->itemNew.lParam);
	return 0;
}

LRESULT CCustomTreeDlg::OnAddItem( int idCtrl, LPNMHDR pnmh, BOOL& bHandled )
{
	if (pnmh->code==TVN_KEYDOWN)
	{
		NMTVKEYDOWN *pKey=(NMTVKEYDOWN*)pnmh;
		if (pKey->wVKey!=VK_RETURN)
		{
			bHandled=FALSE;
			return 0;
		}
	}

	HTREEITEM hItem=TreeView_GetSelection(m_CommandsTree);
	if (hItem)
		AddItem(hItem);
	return TRUE;
}

LRESULT CCustomTreeDlg::OnEditItem( int idCtrl, LPNMHDR pnmh, BOOL& bHandled )
{
	if (pnmh->code==TVN_KEYDOWN)
	{
		NMTVKEYDOWN *pKey=(NMTVKEYDOWN*)pnmh;
		if (pKey->wVKey!=VK_RETURN)
		{
			bHandled=FALSE;
			return 0;
		}
	}

	TVITEM item={TVIF_PARAM,TreeView_GetSelection(m_Tree)};
	if (!item.hItem) return 0;
	TreeView_GetItem(m_Tree,&item);
	CTreeItem *pItem=(CTreeItem*)item.lParam;
	if (pItem)
		EditItemInternal(pItem,item.hItem);
	return TRUE;
}

void CCustomTreeDlg::AddItem( HTREEITEM hCommand )
{
	TVITEM item={TVIF_PARAM,hCommand};
	TreeView_GetItem(m_CommandsTree,&item);
	const CStdCommand *pCommand=(CStdCommand*)item.lParam;
	HTREEITEM hAfter=TreeView_GetSelection(m_Tree);
	HTREEITEM hParent=NULL;
	if (hAfter)
	{
		hParent=TreeView_GetParent(m_Tree,hAfter);
		TVITEM item={TVIF_PARAM,hAfter};
		TreeView_GetItem(m_Tree,&item);
		if (!item.lParam)
		{
			TreeView_DeleteItem(m_Tree,hAfter);
			hAfter=TVI_LAST;
		}
	}
	else
		hAfter=TVI_LAST;
	HTREEITEM hItem=m_Tree.CreateStdItem(pCommand,hParent,hAfter);
	TreeView_SelectItem(m_Tree,hItem);
	TreeView_EnsureVisible(m_Tree,hItem);
	SerializeData();
}

HTREEITEM CCustomTreeDlg::GetRoot( void )
{
	return TreeView_GetRoot(m_Tree);
}

HTREEITEM CCustomTreeDlg::GetChild( HTREEITEM hParent )
{
	return TreeView_GetChild(m_Tree,hParent);
}

HTREEITEM CCustomTreeDlg::GetNext( HTREEITEM hItem )
{
	return TreeView_GetNextSibling(m_Tree,hItem);
}

CTreeItem *CCustomTreeDlg::GetItem( HTREEITEM hItem )
{
	TVITEM item={TVIF_PARAM,hItem};
	TreeView_GetItem(m_Tree,&item);
	return (CTreeItem*)item.lParam;
}

void CCustomTreeDlg::EditItemInternal( CTreeItem *pItem, HTREEITEM hItem )
{
	if (!pItem->bSeparator && EditItem(pItem,m_Tree,hItem,m_Tree.m_Modules))
	{
		SerializeData();
		m_Tree.UpdateIcons(pItem,hItem);
	}
}

bool CCustomTreeDlg::EditItem( CTreeItem *pItem, HWND tree, HTREEITEM hItem, std::vector<HMODULE> &modules )
{
	MessageBox(L"Edit Item",L"",MB_OK);
	return false;
}

LRESULT CCustomTreeDlg::OnCustomDraw( int idCtrl, LPNMHDR pnmh, BOOL& bHandled )
{
	NMTVCUSTOMDRAW *pDraw=(NMTVCUSTOMDRAW*)pnmh;
	if (pDraw->nmcd.dwDrawStage==CDDS_PREPAINT)
	{
		return CDRF_NOTIFYITEMDRAW;
	}
	else if (pDraw->nmcd.dwDrawStage==CDDS_ITEMPREPAINT && !pDraw->nmcd.lItemlParam)
	{
		if (!IsAppThemed())
		{
			TVITEM item={TVIF_STATE,(HTREEITEM)pDraw->nmcd.dwItemSpec,0,TVIS_SELECTED};
			TreeView_GetItem(m_Tree,&item);
			if (item.state&TVIS_SELECTED)
				return CDRF_DODEFAULT;
		}
		pDraw->clrText=GetSysColor(COLOR_GRAYTEXT);
	}
	return CDRF_DODEFAULT;
}

int CCustomTreeDlg::ParseTreeItem( CTreeItem *pItem, CSettingsParser &parser )
{
	pItem->bSeparator=IsSeparator(pItem->name);
	if (pItem->bSeparator) return 0;
	wchar_t text[256];
	Sprintf(text,_countof(text),L"%s.Command",pItem->name);
	pItem->SetCommand(parser.FindSetting(text,L""),m_pStdCommands);
	Sprintf(text,_countof(text),L"%s.Link",pItem->name);
	pItem->link=parser.FindSetting(text,L"");
	Sprintf(text,_countof(text),L"%s.Label",pItem->name);
	pItem->label=parser.FindSetting(text,L"");
	Sprintf(text,_countof(text),L"%s.Tip",pItem->name);
	pItem->tip=parser.FindSetting(text,L"");
	Sprintf(text,_countof(text),L"%s.Icon",pItem->name);
	pItem->icon=parser.FindSetting(text,L"");
	ParseTreeItemExtra(pItem,parser);
	return m_Tree.UpdateIcons(pItem,NULL);
}

void CCustomTreeDlg::AppendString( std::vector<wchar_t> &stringBuilder, const wchar_t *text )
{
	int len=Strlen(text);
	int pos=(int)stringBuilder.size();
	stringBuilder.resize(pos+len);
	memcpy(&stringBuilder[pos],text,len*2);
}

void CCustomTreeDlg::SerializeItem( HTREEITEM hItem, std::vector<wchar_t> &stringBuilder )
{
	TVITEM item={TVIF_PARAM,hItem};
	CTreeItem *pItem=NULL;
	if (hItem)
	{
		TreeView_GetItem(m_Tree,&item);
		pItem=(CTreeItem*)item.lParam;
	}
	wchar_t text[2048];
	HTREEITEM hChild=hItem?TreeView_GetChild(m_Tree,hItem):TreeView_GetRoot(m_Tree);
	if (hChild)
	{
		// save .Items
		if (pItem)
			Sprintf(text,_countof(text),L"%s.Items=",pItem->name);
		else
			Strcpy(text,_countof(text),L"Items=");
		AppendString(stringBuilder,text);

		bool bFirst=true;
		for (;hChild;hChild=TreeView_GetNextSibling(m_Tree,hChild), bFirst=false)
		{
			item.hItem=hChild;
			TreeView_GetItem(m_Tree,&item);
			CTreeItem *pChild=(CTreeItem*)item.lParam;
			if (pChild)
			{
				Sprintf(text,_countof(text),bFirst?L"%s":L",%s",pChild->name);
				AppendString(stringBuilder,text);
			}
		}
		AppendString(stringBuilder,L"\n");
	}

	if (pItem && !pItem->bSeparator)
	{
		// save other properties
		if (!pItem->command.IsEmpty())
		{
			Sprintf(text,_countof(text),L"%s.Command=%s\n",pItem->name,pItem->command);
			AppendString(stringBuilder,text);
		}
		if (!pItem->link.IsEmpty())
		{
			Sprintf(text,_countof(text),L"%s.Link=%s\n",pItem->name,pItem->link);
			AppendString(stringBuilder,text);
		}
		if (!pItem->label.IsEmpty())
		{
			Sprintf(text,_countof(text),L"%s.Label=%s\n",pItem->name,pItem->label);
			AppendString(stringBuilder,text);
		}
		if (!pItem->tip.IsEmpty())
		{
			Sprintf(text,_countof(text),L"%s.Tip=%s\n",pItem->name,pItem->tip);
			AppendString(stringBuilder,text);
		}
		if (!pItem->icon.IsEmpty())
		{
			Sprintf(text,_countof(text),L"%s.Icon=%s\n",pItem->name,pItem->icon);
			AppendString(stringBuilder,text);
		}
		SerializeItemExtra(pItem,stringBuilder);
	}
	hChild=hItem?TreeView_GetChild(m_Tree,hItem):TreeView_GetRoot(m_Tree);
	for (;hChild;hChild=TreeView_GetNextSibling(m_Tree,hChild))
		SerializeItem(hChild,stringBuilder);
}

void CCustomTreeDlg::SerializeData( void )
{
	std::vector<wchar_t> stringBuilder;
	SerializeItem(NULL,stringBuilder);
	stringBuilder.push_back(0);
	CSettingsLockWrite lock;
	m_pSetting->value=CComVariant(&stringBuilder[0]);
	m_pSetting->flags&=~CSetting::FLAG_DEFAULT;
	ItemsChanged();
}

void CCustomTreeDlg::CreateTreeItems( CSettingsParser &parser, HTREEITEM hParent, const CSettingsParser::TreeItem *pItems, int index )
{
	for (;!pItems[index].name.IsEmpty();index++)
	{
		TVINSERTSTRUCT insert={hParent,TVI_LAST,{TVIF_STATE|TVIF_TEXT|TVIF_PARAM|TVIF_IMAGE|TVIF_SELECTEDIMAGE,NULL,TVIS_EXPANDED,TVIS_EXPANDED,(LPWSTR)(LPCWSTR)pItems[index].name}};
		CTreeItem *pNewItem=new CTreeItem;
		pNewItem->name=pItems[index].name;
		insert.item.iImage=insert.item.iSelectedImage=ParseTreeItem(pNewItem,parser);
		insert.item.lParam=(LPARAM)pNewItem;
		HTREEITEM hItem=TreeView_InsertItem(m_Tree,&insert);

		if (pItems[index].children>=0)
			CreateTreeItems(parser,hItem,pItems,pItems[index].children);
	}
}

void CCustomTreeDlg::SetGroup( CSetting *pGroup, bool bReset )
{
	if (!m_pSetting || bReset) // optimization - initialize the UI only the first time
	{
		TreeView_DeleteAllItems(m_Tree);
		m_pSetting=pGroup+1;
		m_Tree.SetDisabled((m_pSetting->flags&CSetting::FLAG_LOCKED_MASK)!=0);
		const wchar_t *pText=L"";
		if (m_pSetting->value.vt==VT_BSTR)
			pText=m_pSetting->value.bstrVal;

		CSettingsParser parser;
		parser.LoadText(pText,Strlen(pText));
		parser.ParseText();
		std::vector<CSettingsParser::TreeItem> items;
		parser.ParseTree(L"Items",items);
		if (!items.empty())
			CreateTreeItems(parser,NULL,&items[0],0);
		InitItems();
	}
	// use the tooltip from m_CommandsTree to show the "locked" tooltip for m_Tree. otherwise it will behave like an infotip and move from item to item
	CWindow tooltip=(HWND)m_CommandsTree.SendMessage(TVM_GETTOOLTIPS);
	TOOLINFO tool={sizeof(tool),TTF_SUBCLASS|TTF_IDISHWND,m_hWnd,(UINT_PTR)(HWND)m_Tree};
	CString str=LoadStringEx(m_bMenu?IDS_MENU_LOCKED:IDS_TOOLBAR_LOCKED);
	tool.lpszText=(LPWSTR)(LPCWSTR)str;
	if (m_pSetting->flags&CSetting::FLAG_LOCKED_REG)
		tooltip.SendMessage(TTM_ADDTOOL,0,(LPARAM)&tool);
	else
		tooltip.SendMessage(TTM_DELTOOL,0,(LPARAM)&tool);
}

CCustomTreeDlg::CCustomTreeDlg( bool bMenu, const CStdCommand *pStdCommands ):
m_Tree(*new CSettingsTree), m_CommandsTree(*new CCommandsTree)
{
	m_bMenu=bMenu;
	m_pStdCommands=pStdCommands;
}

CCustomTreeDlg::~CCustomTreeDlg( void )
{
	delete &m_Tree;
	delete &m_CommandsTree;
}

///////////////////////////////////////////////////////////////////////////////

class CBrowseForIconDlg: public CResizeableDlg<CBrowseForIconDlg>
{
public:
	CBrowseForIconDlg( wchar_t *path, int &id ): m_pPath(path), m_ID(id) { m_bResult=false; }
	void SetEnableParent( HWND parent ) { m_EnableParent=parent; }
	bool GetResult( void ) { return m_bResult; }

	BEGIN_MSG_MAP( CBrowseForIconDlg )
		MESSAGE_HANDLER( WM_INITDIALOG, OnInitDialog )
		MESSAGE_HANDLER( WM_SIZE, OnSize )
		MESSAGE_HANDLER( WM_GETMINMAXINFO, OnGetMinMaxInfo )
		COMMAND_ID_HANDLER( IDC_BUTTONBROWSE, OnBrowse )
		COMMAND_ID_HANDLER( IDOK, OnOK )
		COMMAND_ID_HANDLER( IDCANCEL, OnCancel )
		COMMAND_HANDLER( IDC_EDITFILE, EN_KILLFOCUS, OnKillFocus )
		NOTIFY_HANDLER( IDC_LISTICONS, LVN_ITEMACTIVATE, OnItemActivate )
	END_MSG_MAP()

	BEGIN_RESIZE_MAP
		RESIZE_CONTROL(IDC_EDITFILE,MOVE_SIZE_X)
		RESIZE_CONTROL(IDC_BUTTONBROWSE,MOVE_MOVE_X)
		RESIZE_CONTROL(IDC_LISTICONS,MOVE_SIZE_X|MOVE_SIZE_Y)
		RESIZE_CONTROL(IDOK,MOVE_MOVE_X|MOVE_MOVE_Y)
		RESIZE_CONTROL(IDCANCEL,MOVE_MOVE_X|MOVE_MOVE_Y)
	END_RESIZE_MAP

protected:
	// Handler prototypes:
	//  LRESULT MessageHandler(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
	//  LRESULT CommandHandler(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled);
	//  LRESULT NotifyHandler(int idCtrl, LPNMHDR pnmh, BOOL& bHandled);
	LRESULT OnInitDialog( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled );
	LRESULT OnSize( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled );
	LRESULT OnBrowse( WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled );
	LRESULT OnOK( WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled );
	LRESULT OnCancel( WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled );
	LRESULT OnKillFocus( WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled );
	LRESULT OnItemActivate( int idCtrl, LPNMHDR pnmh, BOOL& bHandled );

private:
	HWND m_EnableParent;
	bool m_bResult;
	wchar_t *m_pPath;
	int &m_ID;
	CString m_LastName;

	void UpdateIcons( int id );
	void StorePlacement( void );
	static BOOL CALLBACK EnumIconsProc( HMODULE hModule, LPCTSTR lpszType, LPTSTR lpszName, LONG_PTR lParam );
};

LRESULT CBrowseForIconDlg::OnInitDialog( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled )
{
	SetDlgItemText(IDC_EDITFILE,m_pPath);
	CWindow list=GetDlgItem(IDC_LISTICONS);
	list.SendMessage(LVM_SETEXTENDEDLISTVIEWSTYLE,LVS_EX_DOUBLEBUFFER,LVS_EX_DOUBLEBUFFER);
	HDC hdc=CreateCompatibleDC(NULL);
	HFONT font0=(HFONT)SelectObject(hdc,list.GetFont());
	TEXTMETRIC metrics;
	GetTextMetrics(hdc,&metrics);
	SelectObject(hdc,font0);
	DeleteDC(hdc);
	ListView_SetIconSpacing(list,GetSystemMetrics(SM_CXICON)+metrics.tmHeight+7,GetSystemMetrics(SM_CYICON)+metrics.tmHeight+7);
	if (IsAppThemed())
		SetWindowTheme(list,L"Explorer",NULL);
	UpdateIcons(m_ID);
	CResizeableDlg<CBrowseForIconDlg>::InitResize(MOVE_MODAL);
	SHAutoComplete(GetDlgItem(IDC_EDITFILE),SHACF_FILESYSTEM);

	RECT rc;
	CRegKey regSettings;
	if (regSettings.Open(HKEY_CURRENT_USER,GetSettingsRegPath())==ERROR_SUCCESS)
	{
		ULONG size=sizeof(rc);
		if (regSettings.QueryBinaryValue(L"CBrowseForIconDlg",&rc,&size)==ERROR_SUCCESS && size==sizeof(rc))
			SetStoreRect(rc);
	}
	return TRUE;
}

LRESULT CBrowseForIconDlg::OnSize( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled )
{
	CResizeableDlg<CBrowseForIconDlg>::OnSize();
	return 0;
}

LRESULT CBrowseForIconDlg::OnBrowse( WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled )
{
	OPENFILENAME ofn={sizeof(ofn)};
	ofn.hwndOwner=m_hWnd;
	wchar_t filters[256];
	Strcpy(filters,_countof(filters),LoadStringEx(IDS_ICON_FILTERS));
	for (wchar_t *c=filters;*c;c++)
		if (*c=='|') *c=0;
	ofn.lpstrFilter=filters;
	ofn.nFilterIndex=1;
	wchar_t path[_MAX_PATH];
	GetDlgItemText(IDC_EDITFILE,path,_countof(path));
	DoEnvironmentSubst(path,_countof(path));
	ofn.lpstrFile=path;
	ofn.nMaxFile=_countof(path);
	CString title=LoadStringEx(IDS_ICON_TITLE);
	ofn.lpstrTitle=title;
	ofn.Flags=OFN_DONTADDTORECENT|OFN_ENABLESIZING|OFN_EXPLORER|OFN_FILEMUSTEXIST|OFN_HIDEREADONLY|OFN_NOCHANGEDIR;
	bool bRes=false;
	if (GetOpenFileName(&ofn))
	{
		bRes=true;
	}
	else if (CommDlgExtendedError()==FNERR_INVALIDFILENAME)
	{
		// bad path. clear path and try again
		path[0]=0;
		if (GetOpenFileName(&ofn))
			bRes=true;
	}
	if (bRes)
	{
		wchar_t buf[_MAX_PATH];
		UnExpandEnvStrings(path,buf,_countof(buf));
		SetDlgItemText(IDC_EDITFILE,buf);
		UpdateIcons(-1);
	}
	return 0;
}

void CBrowseForIconDlg::StorePlacement( void )
{
	RECT rc;
	GetStoreRect(rc);

	CRegKey regSettings;
	if (regSettings.Open(HKEY_CURRENT_USER,GetSettingsRegPath())!=ERROR_SUCCESS)
		regSettings.Create(HKEY_CURRENT_USER,GetSettingsRegPath());
	regSettings.SetBinaryValue(L"CBrowseForIconDlg",&rc,sizeof(rc));
}

LRESULT CBrowseForIconDlg::OnOK( WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled )
{
	if (GetFocus()==GetDlgItem(IDC_EDITFILE))
	{
		SendMessage(WM_NEXTDLGCTL,(WPARAM)GetDlgItem(IDOK).m_hWnd,TRUE);
		return 0;
	}
	CString text;
	GetDlgItemText(IDC_EDITFILE,text);
	text.TrimLeft();
	text.TrimRight();
	Strcpy(m_pPath,_MAX_PATH,text);
	CWindow list=GetDlgItem(IDC_LISTICONS);
	int sel=ListView_GetNextItem(list,-1,LVNI_SELECTED);
	if (sel>=0)
	{
		LVITEM item={LVIF_PARAM,sel};
		ListView_GetItem(list,&item);
		m_ID=(int)item.lParam;
	}
	else
		m_ID=0;
	::EnableWindow(m_EnableParent,TRUE);
	m_bResult=true;
	StorePlacement();
	DestroyWindow();
	return 0;
}

LRESULT CBrowseForIconDlg::OnCancel( WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled )
{
	::EnableWindow(m_EnableParent,TRUE);
	m_bResult=false;
	StorePlacement();
	DestroyWindow();
	return 0;
}

LRESULT CBrowseForIconDlg::OnKillFocus( WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled )
{
	UpdateIcons(-1);
	return 0;
}

LRESULT CBrowseForIconDlg::OnItemActivate( int idCtrl, LPNMHDR pnmh, BOOL& bHandled )
{
	PostMessage(WM_COMMAND,IDOK,0);
	return 0;
}

BOOL CALLBACK CBrowseForIconDlg::EnumIconsProc( HMODULE hModule, LPCTSTR lpszType, LPTSTR lpszName, LONG_PTR lParam )
{
	if (IS_INTRESOURCE(lpszName))
	{
		HWND list=(HWND)lParam;
		HICON hIcon=(HICON)LoadImage(hModule,lpszName,IMAGE_ICON,0,0,LR_DEFAULTSIZE|LR_DEFAULTCOLOR);
		if (hIcon)
		{
			HIMAGELIST images=ListView_GetImageList(list,LVSIL_NORMAL);
			int idx=ImageList_AddIcon(images,hIcon);
			DestroyIcon(hIcon);
			LVITEM item={LVIF_IMAGE|LVIF_PARAM|LVIF_TEXT};
			item.iItem=idx;
			item.iImage=idx;
			item.lParam=(WORD)(DWORD_PTR)lpszName;
			wchar_t text[20];
			Sprintf(text,_countof(text),L"%d",(int)item.lParam);
			item.pszText=text;
			ListView_InsertItem(list,&item);
		}
	}
	return TRUE;
}

void CBrowseForIconDlg::UpdateIcons( int id )
{
	wchar_t path[1024];
	GetDlgItemText(IDC_EDITFILE,path,_countof(path));
	DoEnvironmentSubst(path,_countof(path));
	if (m_LastName==path && id==-1) return;
	m_LastName=path;

	CWindow list=GetDlgItem(IDC_LISTICONS);
	ListView_DeleteAllItems(list);
	HIMAGELIST images=ImageList_Create(GetSystemMetrics(SM_CXICON),GetSystemMetrics(SM_CYICON),ILC_COLOR32|ILC_MASK,10,10);
	ListView_SetImageList(list,images,LVSIL_NORMAL);
	ImageList_RemoveAll(images);

	if (_wcsicmp(PathFindExtension(path),L".ico")==0)
	{
		// load icon file
		HICON hIcon=(HICON)LoadImage(NULL,path,IMAGE_ICON,0,0,LR_DEFAULTSIZE|LR_DEFAULTCOLOR|LR_LOADFROMFILE);
		if (hIcon)
		{
			ImageList_AddIcon(images,hIcon);
			DestroyIcon(hIcon);
			LVITEM item={LVIF_IMAGE|LVIF_PARAM|LVIF_STATE,0,0,LVIS_SELECTED,LVIS_SELECTED};
			item.iImage=0;
			item.lParam=0;
			ListView_InsertItem(list,&item);
		}
	}
	else
	{
		// load module file
		HINSTANCE hInstance=_AtlBaseModule.GetResourceInstance();
		HMODULE hModule;
		if (*path)
			hModule=LoadLibraryEx(path,NULL,LOAD_LIBRARY_AS_DATAFILE|LOAD_LIBRARY_AS_IMAGE_RESOURCE);
		else
			hModule=hInstance;
		if (hModule)
		{
			// enumerate icons, extract icons
			list.SendMessage(WM_SETREDRAW,FALSE);
			EnumResourceNames(hModule,RT_GROUP_ICON,EnumIconsProc,(LONG_PTR)list.m_hWnd);
			list.SendMessage(WM_SETREDRAW,TRUE);
			// free module
			if (hModule!=hInstance) FreeLibrary(hModule);
		}
		if (ListView_GetItemCount(list)>0)
		{
			int idx=0;
			if (id>0)
			{
				LVFINDINFO find={LVFI_PARAM,0,id};
				idx=ListView_FindItem(list,-1,&find);
				if (idx<0) idx=0;
			}
			ListView_SetItemState(list,idx,LVIS_SELECTED|LVIS_FOCUSED,LVIS_SELECTED|LVIS_FOCUSED);
			ListView_EnsureVisible(list,idx,FALSE);
			list.InvalidateRect(NULL);
		}
	}
}

bool BrowseForIcon( HWND hWndParent, wchar_t *path, int &id )
{
	CBrowseForIconDlg dlg(path,id);

	hWndParent=GetAncestor(hWndParent,GA_ROOT);
	DLGTEMPLATE *pTemplate=LoadDialogEx(IDD_BROWSEFORICON);
	dlg.Create(hWndParent,pTemplate);
	dlg.ShowWindow(SW_SHOWNORMAL);
	EnableWindow(hWndParent,FALSE);
	dlg.SetEnableParent(hWndParent);
	MSG msg;
	while (dlg.m_hWnd && GetMessage(&msg,0,0,0))
	{
		if (dlg.IsDialogMessage(&msg)) continue;
		TranslateMessage(&msg);
		DispatchMessage(&msg);
	}
	return dlg.GetResult();
}

///////////////////////////////////////////////////////////////////////////////

static bool BrowseForBitmap( HWND hWndParent, wchar_t *path )
{
	OPENFILENAME ofn={sizeof(ofn)};
	ofn.hwndOwner=hWndParent;
	wchar_t filters[256];
	Strcpy(filters,_countof(filters),LoadStringEx(IDS_BMP_FILTERS));
	for (wchar_t *c=filters;*c;c++)
		if (*c=='|') *c=0;
	ofn.lpstrFilter=filters;
	ofn.nFilterIndex=1;
	ofn.lpstrFile=path;
	ofn.nMaxFile=_MAX_PATH;
	CString title=LoadStringEx(IDS_BMP_TITLE);
	ofn.lpstrTitle=title;
	ofn.Flags=OFN_DONTADDTORECENT|OFN_ENABLESIZING|OFN_EXPLORER|OFN_FILEMUSTEXIST|OFN_HIDEREADONLY|OFN_NOCHANGEDIR;
	bool bRes=false;
	if (GetOpenFileName(&ofn))
	{
		bRes=true;
	}
	else if (CommDlgExtendedError()==FNERR_INVALIDFILENAME)
	{
		// bad path. clear path and try again
		path[0]=0;
		if (GetOpenFileName(&ofn))
			bRes=true;
	}
	if (bRes)
	{
		wchar_t buf[_MAX_PATH];
		UnExpandEnvStrings(path,buf,_countof(buf));
		Strcpy(path,_MAX_PATH,buf);
	}
	return bRes;
}

///////////////////////////////////////////////////////////////////////////////

static bool BrowseForSound( HWND hWndParent, wchar_t *path )
{
	OPENFILENAME ofn={sizeof(ofn)};
	ofn.hwndOwner=hWndParent;
	wchar_t filters[256];
	Strcpy(filters,_countof(filters),LoadStringEx(IDS_WAV_FILTERS));
	for (wchar_t *c=filters;*c;c++)
		if (*c=='|') *c=0;
	ofn.lpstrFilter=filters;
	ofn.nFilterIndex=1;
	ofn.lpstrFile=path;
	ofn.nMaxFile=_MAX_PATH;
	CString title=LoadStringEx(IDS_WAV_TITLE);
	ofn.lpstrTitle=title;
	ofn.Flags=OFN_DONTADDTORECENT|OFN_ENABLESIZING|OFN_EXPLORER|OFN_FILEMUSTEXIST|OFN_HIDEREADONLY|OFN_NOCHANGEDIR;
	bool bRes=false;
	if (GetOpenFileName(&ofn))
	{
		bRes=true;
	}
	else if (CommDlgExtendedError()==FNERR_INVALIDFILENAME)
	{
		// bad path. clear path and try again
		path[0]=0;
		if (GetOpenFileName(&ofn))
			bRes=true;
	}
	if (bRes)
	{
		wchar_t buf[_MAX_PATH];
		UnExpandEnvStrings(path,buf,_countof(buf));
		Strcpy(path,_MAX_PATH,buf);
	}
	return bRes;
}

///////////////////////////////////////////////////////////////////////////////

class CTreeSettingsDlg: public CResizeableDlg<CTreeSettingsDlg>
{
public:
	enum {
		IDC_EDITBOX=101,
		IDC_BROWSEBTN=102,
		IDC_PLAYBTN=103,
		WM_EDITKEY=WM_APP+11,
	};

	enum TEditMode
	{
		EDIT_NONE,
		EDIT_INT,
		EDIT_STRING,
		EDIT_ICON,
		EDIT_BITMAP,
		EDIT_SOUND,
		EDIT_HOTKEY,
		EDIT_HOTKEY_ANY,
		EDIT_COLOR,
		EDIT_FONT,
	};

	BEGIN_MSG_MAP( CTreeSettingsDlg )
		MESSAGE_HANDLER( WM_INITDIALOG, OnInitDialog )
		MESSAGE_HANDLER( WM_DESTROY, OnDestroy )
		MESSAGE_HANDLER( WM_SIZE, OnSize )
		MESSAGE_HANDLER( WM_EDITKEY, OnEditKey )
		MESSAGE_HANDLER( WM_CONTEXTMENU, OnContextMenu )
		COMMAND_HANDLER( IDC_EDITBOX, EN_CHANGE, OnEditChange )
		COMMAND_HANDLER( IDC_EDITBOX, EN_KILLFOCUS, OnKillFocus )
		COMMAND_HANDLER( IDC_BROWSEBTN, BN_CLICKED, OnBrowse )
		COMMAND_HANDLER( IDC_PLAYBTN, BN_CLICKED, OnPlay )
		COMMAND_HANDLER( IDC_BROWSEBTN, BN_KILLFOCUS, OnKillFocus )
		NOTIFY_HANDLER( IDC_SETTINGS, NM_CUSTOMDRAW, OnCustomDraw )
		NOTIFY_HANDLER( IDC_SETTINGS, NM_CLICK, OnClick )
		NOTIFY_HANDLER( IDC_SETTINGS, NM_DBLCLK, OnClick )
		NOTIFY_HANDLER( IDC_SETTINGS, TVN_KEYDOWN, OnKeyDown )
		NOTIFY_HANDLER( IDC_SETTINGS, TVN_GETINFOTIP, OnGetInfoTip )
		NOTIFY_HANDLER( IDC_SETTINGS, TVN_SELCHANGED, OnSelChanged )
	END_MSG_MAP()

	BEGIN_RESIZE_MAP
		RESIZE_CONTROL(IDC_SETTINGS,MOVE_SIZE_X|MOVE_SIZE_Y)
		RESIZE_CONTROL(IDC_STATICTIP,MOVE_SIZE_X|MOVE_MOVE_Y)
	END_RESIZE_MAP

	void SetGroup( CSetting *pGroup );
	void Validate( void ) { ApplyEditBox(); }

protected:
	// Handler prototypes:
	//  LRESULT MessageHandler(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
	//  LRESULT CommandHandler(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled);
	//  LRESULT NotifyHandler(int idCtrl, LPNMHDR pnmh, BOOL& bHandled);
	LRESULT OnInitDialog( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled );
	LRESULT OnDestroy( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled );
	LRESULT OnSize( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled );
	LRESULT OnEditKey( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled );
	LRESULT OnContextMenu( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled );
	LRESULT OnEditChange( WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled );
	LRESULT OnKillFocus( WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled );
	LRESULT OnBrowse( WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled );
	LRESULT OnPlay( WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled );
	LRESULT OnCustomDraw( int idCtrl, LPNMHDR pnmh, BOOL& bHandled );
	LRESULT OnKeyDown( int idCtrl, LPNMHDR pnmh, BOOL& bHandled );
	LRESULT OnClick( int idCtrl, LPNMHDR pnmh, BOOL& bHandled );
	LRESULT OnGetInfoTip( int idCtrl, LPNMHDR pnmh, BOOL& bHandled );
	LRESULT OnSelChanged( int idCtrl, LPNMHDR pnmh, BOOL& bHandled );

private:
	CWindow m_Tree;
	const CSetting *m_pGroup;
	CWindow m_EditBox;
	CWindow m_BrowseButton;
	CWindow m_PlayButton;
	HICON m_PlayIcon;
	TEditMode m_EditMode;
	HTREEITEM m_EditItem;
	CSetting *m_pEditSetting;
	bool m_bBasic;
	bool m_bIgnoreFocus;

	void UpdateGroup( const CSetting *pModified );
	void ToggleItem( HTREEITEM hItem, bool bDefault );
	void PlayItem( HTREEITEM hItem );
	void ItemSelected( HTREEITEM hItem, CSetting *pSetting, bool bEnabled );
	void ApplyEditBox( void );
	void UpdateEditPosition( void );
	const CSetting *GetNextSetting( const CSetting *pSetting );
};

// Subclass the tree control to forward the command messages to the parent
static LRESULT CALLBACK SubclassTreeProc( HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData )
{
	if (uMsg==WM_COMMAND)
		return SendMessage(GetParent(hWnd),uMsg,wParam,lParam);
	return DefSubclassProc(hWnd,uMsg,wParam,lParam);
}

static int g_HotKey;

static bool GetKeyName( int key, wchar_t *text, int size, bool bCtrl, bool bAlt, bool bShift )
{
	wchar_t name[100];
	text[0]=0;
	int len=0;
	if (bCtrl || (key&(HOTKEYF_CONTROL<<8)))
	{
		GetKeyNameText(MapVirtualKey(VK_CONTROL,MAPVK_VK_TO_VSC)<<16,name,_countof(name));
		len+=Sprintf(text+len,size-len,L"%s+",name);
	}
	if (bAlt || (key&(HOTKEYF_ALT<<8)))
	{
		GetKeyNameText(MapVirtualKey(VK_MENU,MAPVK_VK_TO_VSC)<<16,name,_countof(name));
		len+=Sprintf(text+len,size-len,L"%s+",name);
	}
	if (bShift || (key&(HOTKEYF_SHIFT<<8)))
	{
		GetKeyNameText(MapVirtualKey(VK_SHIFT,MAPVK_VK_TO_VSC)<<16,name,_countof(name));
		len+=Sprintf(text+len,size-len,L"%s+",name);
	}
	if (key)
	{
		int scan=MapVirtualKey(key&255,MAPVK_VK_TO_VSC);
		if (scan==0)
		{
			text[0]=0;
			return false;
		}
		if (!GetKeyNameText(scan<<16,name,_countof(name)))
		{
			text[0]=0;
			return false;
		}
		Strcpy(text+len,size-len,name);
	}

	return true;
}

// Subclass the edit control to forward some keyboard messages to the parent and to act as a hotkey control
static LRESULT CALLBACK SubclassEditProc( HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData )
{
	if (dwRefData)
	{
		static bool s_bCtrl, s_bAlt, s_bShift;

		if (uMsg==WM_GETDLGCODE)
			return DLGC_WANTALLKEYS;

		if (uMsg==WM_CHAR || uMsg==WM_SYSCHAR)
			return 0;

		if (uMsg==WM_KEYDOWN || uMsg==WM_SYSKEYDOWN)
		{
			if ((wParam==VK_UP || wParam==VK_DOWN || wParam==VK_PRIOR || wParam==VK_NEXT) && GetKeyState(VK_CONTROL)>=0 && GetKeyState(VK_MENU)>=0 && GetKeyState(VK_SHIFT)>=0)
				return SendMessage(GetParent(hWnd),uMsg,wParam,lParam);
			if (wParam==VK_ESCAPE)
			{ // Esc - clears the key
				g_HotKey=0;
				SetWindowText(hWnd,L"");
				return 0;
			}
			wchar_t text[100];
			if (wParam==VK_CONTROL) s_bCtrl=true, g_HotKey=0;
			else if (wParam==VK_SHIFT) s_bShift=true, g_HotKey=0;
			else if (wParam==VK_MENU) s_bAlt=true, g_HotKey=0;
			else
			{
				g_HotKey=wParam&255;
				if (s_bCtrl) g_HotKey|=(HOTKEYF_CONTROL<<8);
				if (s_bAlt) g_HotKey|=(HOTKEYF_ALT<<8);
				if (s_bShift) g_HotKey|=(HOTKEYF_SHIFT<<8);
				if (dwRefData==1)
				{
					if (!s_bAlt && !(s_bCtrl && s_bShift))
						g_HotKey=0;
				}
			}
			GetKeyName(g_HotKey,text,_countof(text),s_bCtrl,s_bAlt,s_bShift);
			SetWindowText(hWnd,text);
			return 0;
		}
		if (uMsg==WM_KEYUP || uMsg==WM_SYSKEYUP)
		{
			wchar_t text[100];
			if (wParam==VK_CONTROL) s_bCtrl=false;
			if (wParam==VK_MENU) s_bAlt=false;
			if (wParam==VK_SHIFT) s_bShift=false;
			GetKeyName(g_HotKey,text,_countof(text),s_bCtrl,s_bAlt,s_bShift);
			SetWindowText(hWnd,text);
			return 0;
		}
		if (uMsg==WM_RBUTTONUP || uMsg==WM_CONTEXTMENU)
			return 0;
	}
	if (uMsg==WM_CHAR && wParam==VK_SPACE && GetKeyState(VK_CONTROL)<0)
		return SendMessage(GetParent(hWnd),WM_KEYDOWN,wParam,lParam);
	if (uMsg==WM_CHAR && wParam==VK_RETURN)
		return 0;
	if (uMsg==WM_GETDLGCODE && wParam==VK_RETURN)
		return DLGC_WANTALLKEYS;
	if (uMsg==WM_KEYDOWN)
	{
		if (wParam==VK_UP || wParam==VK_DOWN || wParam==VK_PRIOR || wParam==VK_NEXT || (wParam=='P' && GetKeyState(VK_CONTROL)<0))
			return SendMessage(GetParent(hWnd),uMsg,wParam,lParam);
		if (wParam==VK_RETURN || wParam==VK_ESCAPE)
			return SendMessage(GetParent(GetParent(hWnd)),CTreeSettingsDlg::WM_EDITKEY,wParam,0);
	}
	if (uMsg==WM_KILLFOCUS)
	{
		if (wParam && (HWND)wParam==GetDlgItem(GetParent(hWnd),CTreeSettingsDlg::IDC_BROWSEBTN))
			return 0;
	}
	if (uMsg==WM_CHAR && wParam==16)
		return 0; // prevent Ctrl+P from playing the default sound
	return DefSubclassProc(hWnd,uMsg,wParam,lParam);
}

LRESULT CTreeSettingsDlg::OnInitDialog( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled )
{
	CResizeableDlg<CTreeSettingsDlg>::InitResize();
	EnableThemeDialogTexture(m_hWnd,ETDT_ENABLETAB);

	m_Tree=GetDlgItem(IDC_SETTINGS);
	m_Tree.SetWindowLong(GWL_STYLE,m_Tree.GetWindowLong(GWL_STYLE)|WS_CLIPCHILDREN);
	if (IsAppThemed())
	{
		m_Tree.SetWindowLong(GWL_STYLE,m_Tree.GetWindowLong(GWL_STYLE)|TVS_TRACKSELECT);
		SetWindowTheme(m_Tree,L"Explorer",NULL);
	}
	m_Tree.SendMessage(TVM_SETEXTENDEDSTYLE,TVS_EX_DOUBLEBUFFER,TVS_EX_DOUBLEBUFFER);
	SetWindowSubclass(m_Tree,SubclassTreeProc,'CLSH',0);
	HINSTANCE hInstance=_AtlBaseModule.GetResourceInstance();

	HIMAGELIST images=GetSettingsImageList(m_Tree);
	TreeView_SetImageList(m_Tree,images,TVSIL_NORMAL);
	int height=TreeView_GetItemHeight(m_Tree);
	if (height<18) TreeView_SetItemHeight(m_Tree,18);

	HWND tooltip=TreeView_GetToolTips(m_Tree);
	SendMessage(tooltip,TTM_SETDELAYTIME,TTDT_AUTOPOP,10000);
	SendMessage(tooltip,TTM_SETDELAYTIME,TTDT_INITIAL,1000);
	SendMessage(tooltip,TTM_SETDELAYTIME,TTDT_RESHOW,1000);
	SetWindowSubclass(tooltip,SubclassInfoTipProc,'CLSH',0);

	HFONT font=m_Tree.GetFont();
	m_EditBox=CreateWindow(L"EDIT",NULL,WS_CHILD|WS_BORDER|ES_AUTOHSCROLL,0,0,0,0,m_Tree,(HMENU)IDC_EDITBOX,hInstance,NULL);
	m_EditBox.SetFont(font);
	SetWindowSubclass(m_EditBox,SubclassEditProc,'CLSH',0);
	m_BrowseButton=CreateWindow(L"BUTTON",L"...",WS_CHILD|BS_PUSHBUTTON,0,0,0,0,m_Tree,(HMENU)IDC_BROWSEBTN,hInstance,NULL);
	m_BrowseButton.SetFont(font);
	m_PlayIcon=(HICON)LoadImage(hInstance,MAKEINTRESOURCE(IDI_ICONPLAY),IMAGE_ICON,0,0,0);
	m_PlayButton=CreateWindow(L"BUTTON",L">",WS_CHILD|BS_PUSHBUTTON|BS_ICON,0,0,0,0,m_Tree,(HMENU)IDC_PLAYBTN,hInstance,NULL);
	m_PlayButton.SendMessage(BM_SETIMAGE,IMAGE_ICON,(LPARAM)m_PlayIcon);
	m_EditMode=EDIT_NONE;
	m_EditItem=NULL;
	m_pEditSetting=NULL;
	m_bIgnoreFocus=false;

	return TRUE;
}

LRESULT CTreeSettingsDlg::OnDestroy( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled )
{
	DestroyIcon(m_PlayIcon);
	bHandled=FALSE;
	m_EditMode=EDIT_NONE;
	return 0;
}

LRESULT CTreeSettingsDlg::OnSize( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled )
{
	CResizeableDlg<CTreeSettingsDlg>::OnSize();
	UpdateEditPosition();
	return 0;
}

LRESULT CTreeSettingsDlg::OnEditKey( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled )
{
	if (wParam==VK_RETURN && !(m_EditBox.GetWindowLong(GWL_STYLE)&ES_READONLY))
		ItemSelected(m_EditItem,m_pEditSetting,true);
	return 0;
}

LRESULT CTreeSettingsDlg::OnEditChange( WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled )
{
	UpdateEditPosition();
	return 0;
}

LRESULT CTreeSettingsDlg::OnKillFocus( WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled )
{
	if (!m_bIgnoreFocus && m_pEditSetting)
	{
		ApplyEditBox();
		UpdateEditPosition();
	}
	return 0;
}

LRESULT CTreeSettingsDlg::OnBrowse( WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled )
{
	if (m_EditMode==EDIT_ICON)
	{
		m_bIgnoreFocus=true;
		CString str;
		m_EditBox.GetWindowText(str);
		str.TrimLeft(); str.TrimRight();
		wchar_t text[_MAX_PATH*2];
		Strcpy(text,_countof(text),str);
		int id=0;
		if (*text)
		{
			wchar_t *c=wcsrchr(text,',');
			if (c)
			{
				*c=0;
				id=_wtol(c+1);
			}
		}
		else
			Strcpy(text,_countof(text),L"shell32.dll");
		if (BrowseForIcon(m_hWnd,text,id))
		{
			if (id>0)
			{
				wchar_t buf[20];
				Sprintf(buf,_countof(buf),L", %d",id);
				Strcat(text,_countof(text),buf);
			}
			m_EditBox.SetWindowText(text);
		}
		SendMessage(WM_NEXTDLGCTL,(LPARAM)m_EditBox.m_hWnd,TRUE);
		m_bIgnoreFocus=false;
	}
	else if (m_EditMode==EDIT_BITMAP)
	{
		m_bIgnoreFocus=true;
		CString str;
		m_EditBox.GetWindowText(str);
		str.TrimLeft(); str.TrimRight();
		wchar_t text[1024];
		if (_wcsicmp(PathFindExtension(str),L".bmp")==0)
		{
			Strcpy(text,_countof(text),str);
			DoEnvironmentSubst(text,_countof(text));
		}
		else
			text[0]=0;
		if (BrowseForBitmap(m_hWnd,text))
		{
			m_EditBox.SetWindowText(text);
		}
		SendMessage(WM_NEXTDLGCTL,(LPARAM)m_EditBox.m_hWnd,TRUE);
		m_EditBox.SetFocus();
		m_bIgnoreFocus=false;
	}
	else if (m_EditMode==EDIT_SOUND)
	{
		m_bIgnoreFocus=true;
		CString str;
		m_EditBox.GetWindowText(str);
		str.TrimLeft(); str.TrimRight();
		wchar_t text[1024];
		if (_wcsicmp(PathFindExtension(str),L".wav")==0)
		{
			Strcpy(text,_countof(text),str);
			DoEnvironmentSubst(text,_countof(text));
		}
		else
			text[0]=0;
		if (BrowseForSound(m_hWnd,text))
		{
			m_EditBox.SetWindowText(text);
		}
		SendMessage(WM_NEXTDLGCTL,(LPARAM)m_EditBox.m_hWnd,TRUE);
		m_EditBox.SetFocus();
		m_bIgnoreFocus=false;
	}
	else if (m_EditMode==EDIT_COLOR)
	{
		m_bIgnoreFocus=true;
		CString str;
		m_EditBox.GetWindowText(str);
		str.TrimLeft(); str.TrimRight();
		wchar_t *end;
		int val=wcstol(str,&end,16)&0xFFFFFF;
		static COLORREF customColors[16];
		CHOOSECOLOR choose={sizeof(choose),m_hWnd,NULL,val,customColors};
		choose.Flags=CC_ANYCOLOR|CC_FULLOPEN|CC_RGBINIT;
		if (ChooseColor(&choose))
		{
			wchar_t text[100];
			Sprintf(text,_countof(text),L"%06X",choose.rgbResult);
			m_EditBox.SetWindowText(text);
			ApplyEditBox();
			UpdateGroup(m_pEditSetting);
			m_Tree.Invalidate();
		}
		SendMessage(WM_NEXTDLGCTL,(LPARAM)m_EditBox.m_hWnd,TRUE);
		m_EditBox.SetFocus();
		m_bIgnoreFocus=false;
	}
	else if (m_EditMode==EDIT_FONT)
	{
		m_bIgnoreFocus=true;
		CString text;
		m_EditBox.GetWindowText(text);

		HDC hdc=::GetDC(NULL);
		int dpi=GetDeviceCaps(hdc,LOGPIXELSY);
		::ReleaseDC(NULL,hdc);

		LOGFONT font={0};
		const wchar_t *str=text;
		while (*str==' ')
			str++;
		str=GetToken(str,font.lfFaceName,_countof(font.lfFaceName),L",");
		int len=Strlen(font.lfFaceName);
		while (len>0 && font.lfFaceName[len-1]==' ')
			font.lfFaceName[--len]=0;
		while (*str==' ')
			str++;
		wchar_t token[256];
		str=GetToken(str,token,_countof(token),L",");
		len=Strlen(token);
		while (len>0 && token[len-1]==' ')
			token[--len]=0;
		font.lfWeight=FW_NORMAL;
		if (_wcsicmp(token,L"bold")==0)
			font.lfWeight=FW_BOLD;
		else if (_wcsicmp(token,L"italic")==0)
			font.lfItalic=1;
		else if (_wcsicmp(token,L"bold_italic")==0)
			font.lfWeight=FW_BOLD, font.lfItalic=1;
		str=GetToken(str,token,_countof(token),L", \t");
		font.lfHeight=-(_wtol(token)*dpi+36)/72;

		CHOOSEFONT choose={sizeof(choose),m_hWnd,NULL,&font};
		choose.Flags=CF_NOSCRIPTSEL;
		if (*font.lfFaceName)
			choose.Flags|=CF_INITTOLOGFONTSTRUCT;
		if (ChooseFont(&choose))
		{
			wchar_t text[256];
			const wchar_t *type=font.lfItalic?L"italic":L"normal";
			if (font.lfWeight>=FW_BOLD)
				type=font.lfItalic?L"bold_italic":L"bold";
			Sprintf(text,_countof(text),L"%s, %s, %d",font.lfFaceName,type,(-font.lfHeight*72+dpi/2)/dpi);
			m_EditBox.SetWindowText(text);
		}
		SendMessage(WM_NEXTDLGCTL,(LPARAM)m_EditBox.m_hWnd,TRUE);
		m_EditBox.SetFocus();
		m_bIgnoreFocus=false;
	}
	return 0;
}

LRESULT CTreeSettingsDlg::OnPlay( WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled )
{
	PlayItem(TreeView_GetSelection(m_Tree));
	return 0;
}

LRESULT CTreeSettingsDlg::OnCustomDraw( int idCtrl, LPNMHDR pnmh, BOOL& bHandled )
{
	NMTVCUSTOMDRAW *pDraw=(NMTVCUSTOMDRAW*)pnmh;
	if (pDraw->nmcd.dwDrawStage==CDDS_PREPAINT)
	{
		UpdateEditPosition();
		return CDRF_NOTIFYITEMDRAW;
	}
	else if (pDraw->nmcd.dwDrawStage==CDDS_ITEMPREPAINT)
	{
		TVITEM item={TVIF_IMAGE|TVIF_STATE,(HTREEITEM)pDraw->nmcd.dwItemSpec,0,TVIS_SELECTED};
		TreeView_GetItem(m_Tree,&item);
		if ((item.state&TVIS_CUT) && (!(item.state&TVIS_SELECTED) || IsAppThemed()))
			pDraw->clrText=GetSysColor(COLOR_GRAYTEXT);
	}
	return CDRF_DODEFAULT;
}

void CTreeSettingsDlg::ToggleItem( HTREEITEM hItem, bool bDefault )
{
	if (!hItem) return;
	TVITEM item={TVIF_PARAM|TVIF_IMAGE|TVIF_STATE,hItem,0,TVIS_CUT};
	TreeView_GetItem(m_Tree,&item);
	if (item.state&TVIS_CUT)
		return;
	CSetting *pSetting=(CSetting*)item.lParam;
	if (bDefault)
	{
		if (pSetting->flags&CSetting::FLAG_DEFAULT)
			return;
		{
			CSettingsLockWrite lock;
			pSetting->flags|=CSetting::FLAG_DEFAULT;
			pSetting->value=pSetting->defValue;
		}
		if (m_pEditSetting==pSetting)
		{
			m_pEditSetting=NULL;
			ItemSelected(item.hItem,pSetting,true);
		}
		UpdateGroup(pSetting);
	}
	else if (pSetting->type==CSetting::TYPE_BOOL)
	{
		// toggle
		bool state=IsVariantTrue(pSetting->value);
		{
			CSettingsLockWrite lock;
			pSetting->value=CComVariant(state?0:1);
			pSetting->flags&=~CSetting::FLAG_DEFAULT;
		}
		UpdateGroup(pSetting);
	}
	else if (pSetting->type==CSetting::TYPE_RADIO)
	{
		// set parent's value
		for (CSetting *pParent=pSetting-1;;pParent--)
		{
			if (pParent->type!=CSetting::TYPE_RADIO)
			{
				ATLASSERT(pParent->type==CSetting::TYPE_INT);
				int val=(int)(pSetting-pParent-1);
				if (pParent->value.vt==VT_I4 && pParent->value.intVal==val)
					break;
				{
					CSettingsLockWrite lock;
					pParent->value=CComVariant(val);
					pParent->flags&=~CSetting::FLAG_DEFAULT;
				}
				UpdateGroup(pParent);
				break;
			}
		}
	}
}

void CTreeSettingsDlg::PlayItem( HTREEITEM hItem )
{
	if (!hItem) return;
	TVITEM item={TVIF_PARAM,hItem};
	TreeView_GetItem(m_Tree,&item);
	CSetting *pSetting=(CSetting*)item.lParam;
	if (pSetting->type==CSetting::TYPE_SOUND && pSetting->value.vt==VT_BSTR)
	{
		if (!*pSetting->value.bstrVal)
			return;
		wchar_t path[_MAX_PATH];
		Strcpy(path,_countof(path),pSetting->value.bstrVal);
		DoEnvironmentSubst(path,_countof(path));
		if (_wcsicmp(PathFindExtension(path),L"wav")==0)
			PlaySound(path,NULL,SND_FILENAME|SND_ASYNC|SND_NODEFAULT|SND_SYSTEM);
		else
			PlaySound(path,NULL,SND_APPLICATION|SND_ALIAS|SND_ASYNC|SND_NODEFAULT|SND_SYSTEM);
	}
}

LRESULT CTreeSettingsDlg::OnClick( int idCtrl, LPNMHDR pnmh, BOOL& bHandled )
{
	DWORD pos=GetMessagePos();
	TVHITTESTINFO test={{(short)LOWORD(pos),(short)HIWORD(pos)}};
	m_Tree.ScreenToClient(&test.pt);
	if (TreeView_HitTest(m_Tree,&test) && (test.flags&(TVHT_ONITEMICON|TVHT_ONITEMLABEL)))
		ToggleItem(test.hItem,GetKeyState(VK_CONTROL)<0);

	return 0;
}

LRESULT CTreeSettingsDlg::OnContextMenu( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled )
{
	if ((HWND)wParam!=m_Tree.m_hWnd)
		return 0;
	HTREEITEM hItem=NULL;
	POINT pt;
	if ((DWORD)lParam==0xFFFFFFFF)
	{
		// keyboard
		hItem=TreeView_GetSelection(m_Tree);
		if (hItem)
		{
			RECT rc;
			TreeView_GetItemRect(m_Tree,hItem,&rc,TRUE);
			pt.x=(rc.left+rc.right)/2;
			pt.y=(rc.top+rc.bottom)/2;
		}
		else
			pt.x=pt.y=0;
		m_Tree.ClientToScreen(&pt);
	}
	else
	{
		DWORD pos=GetMessagePos();
		pt.x=(short)LOWORD(pos);
		pt.y=(short)HIWORD(pos);
		TVHITTESTINFO test={pt};
		m_Tree.ScreenToClient(&test.pt);
		if (TreeView_HitTest(m_Tree,&test) && (test.flags&(TVHT_ONITEMICON|TVHT_ONITEMLABEL)))
		{
			TreeView_SelectItem(m_Tree,test.hItem);
			hItem=test.hItem;
		}
	}

	if (!hItem) return 0;

	TVITEM item={TVIF_PARAM|TVIF_IMAGE|TVIF_STATE,hItem,0,TVIS_CUT};
	TreeView_GetItem(m_Tree,&item);
	CSetting *pSetting=(CSetting*)item.lParam;

	HMENU menu=CreatePopupMenu();
	if (pSetting->type!=CSetting::TYPE_RADIO)
	{
		if (pSetting->type==CSetting::TYPE_BOOL)
		{
			AppendMenu(menu,MF_STRING,1,LoadStringEx(IDS_TOGGLE_SETTING));
			if (item.state&TVIS_CUT)
				EnableMenuItem(menu,1,MF_BYCOMMAND|MF_GRAYED);
		}
		AppendMenu(menu,MF_STRING,2,LoadStringEx(IDS_DEFAULT_SETTING));
		if ((item.state&TVIS_CUT) || (pSetting->flags&CSetting::FLAG_DEFAULT))
			EnableMenuItem(menu,2,MF_BYCOMMAND|MF_GRAYED);
	}
	else
	{
		AppendMenu(menu,MF_STRING,1,LoadStringEx(IDS_SELECT_SETTING));
		if (item.state&TVIS_CUT)
			EnableMenuItem(menu,1,MF_BYCOMMAND|MF_GRAYED);
	}
	if (pSetting->type==CSetting::TYPE_SOUND)
	{
		AppendMenu(menu,MF_STRING,3,LoadStringEx(IDS_PLAY_SETTING));
		if (pSetting->value.vt!=VT_BSTR || *pSetting->value.bstrVal==0)
			EnableMenuItem(menu,3,MF_BYCOMMAND|MF_GRAYED);
	}

	int res=TrackPopupMenu(menu,TPM_RETURNCMD,pt.x,pt.y,0,m_hWnd,NULL);
	DestroyMenu(menu);
	if (res==1)
		ToggleItem(hItem,false);
	if (res==2)
		ToggleItem(hItem,true);
	if (res==3)
		PlayItem(hItem);

	return 0;
}

LRESULT CTreeSettingsDlg::OnKeyDown( int idCtrl, LPNMHDR pnmh, BOOL& bHandled )
{
	NMTVKEYDOWN *pKey=(NMTVKEYDOWN*)pnmh;
	if (pKey->wVKey==VK_SPACE)
	{
		ToggleItem(TreeView_GetSelection(m_Tree),(GetKeyState(VK_CONTROL)<0));
		return 1;
	}
	if (pKey->wVKey=='P' && GetKeyState(VK_CONTROL)<0)
	{
		PlayItem(TreeView_GetSelection(m_Tree));
		return 1;
	}
	bHandled=FALSE;
	return 0;
}

LRESULT CTreeSettingsDlg::OnGetInfoTip( int idCtrl, LPNMHDR pnmh, BOOL& bHandled )
{
	NMTVGETINFOTIP *pTip=(NMTVGETINFOTIP*)pnmh;
	const CSetting *pSetting=(CSetting*)pTip->lParam;
	if (pSetting->flags&CSetting::FLAG_LOCKED_MASK)
		Sprintf(pTip->pszText,pTip->cchTextMax,L"%s\n%s",LoadStringEx(pSetting->tipID),LoadStringEx((pSetting->flags&CSetting::FLAG_LOCKED_REG)?IDS_SETTING_LOCKED:IDS_SETTING_LOCKED_GP));
	else
		Strcpy(pTip->pszText,pTip->cchTextMax,LoadStringEx(pSetting->tipID));
	return 0;
}

LRESULT CTreeSettingsDlg::OnSelChanged( int idCtrl, LPNMHDR pnmh, BOOL& bHandled )
{
	HTREEITEM hItem=TreeView_GetSelection(m_Tree);
	if (hItem)
	{
		TVITEM item={TVIF_PARAM|TVIF_STATE,hItem,0,TVIS_CUT};
		TreeView_GetItem(m_Tree,&item);
		CSetting *pSetting=(CSetting*)item.lParam;
		ItemSelected(item.hItem,pSetting,(item.state&TVIS_CUT)==0);
	}
	else
		ItemSelected(NULL,NULL,false);
	return 0;
}

void CTreeSettingsDlg::ApplyEditBox( void )
{
	if (m_pEditSetting)
	{
		CSettingsLockWrite lock;

		CString str;
		m_EditBox.GetWindowText(str);
		str.TrimLeft(); str.TrimRight();
		if (m_pEditSetting->type==CSetting::TYPE_INT)
		{
			int val=_wtol(str);
			if (m_pEditSetting->value.vt!=VT_I4 || m_pEditSetting->value.intVal!=val)
			{
				m_pEditSetting->value=CComVariant(val);
				m_pEditSetting->flags&=~CSetting::FLAG_DEFAULT;
			}
		}
		else if (m_pEditSetting->type==CSetting::TYPE_COLOR)
		{
			wchar_t *end;
			int val=wcstol(str,&end,16)&0xFFFFFF;
			if (m_pEditSetting->value.vt!=VT_I4 || m_pEditSetting->value.intVal!=val)
			{
				m_pEditSetting->value=CComVariant(val);
				m_pEditSetting->flags&=~CSetting::FLAG_DEFAULT;
			}
		}
		else if (m_pEditSetting->type==CSetting::TYPE_HOTKEY || m_pEditSetting->type==CSetting::TYPE_HOTKEY_ANY)
		{
			if (m_pEditSetting->value.vt!=VT_I4 || m_pEditSetting->value.intVal!=g_HotKey)
			{
				m_pEditSetting->value=CComVariant(g_HotKey);
				m_pEditSetting->flags&=~CSetting::FLAG_DEFAULT;
			}
		}
		else
		{
			if (m_pEditSetting->value.vt!=VT_BSTR || str!=m_pEditSetting->value.bstrVal)
			{
				m_pEditSetting->value=CComVariant(str);
				m_pEditSetting->flags&=~CSetting::FLAG_DEFAULT;
			}
		}
	}
}

void CTreeSettingsDlg::ItemSelected( HTREEITEM hItem, CSetting *pSetting, bool bEnabled )
{
	if (m_pEditSetting)
	{
		ApplyEditBox();
		const CSetting *pOld=m_pEditSetting;
		m_pEditSetting=NULL;
		UpdateGroup(pOld);
	}
	m_EditItem=NULL;
	TEditMode mode=EDIT_NONE;
	wchar_t text[2048];
	if (pSetting)
	{
		if (pSetting->type==CSetting::TYPE_INT && pSetting[1].type!=CSetting::TYPE_RADIO)
		{
			mode=EDIT_INT;
			int val=0;
			if (pSetting->value.vt==VT_I4)
				val=pSetting->value.intVal;
			Sprintf(text,_countof(text),L"%d",val);
		}
		else if (pSetting->type==CSetting::TYPE_STRING || pSetting->type==CSetting::TYPE_ICON || pSetting->type==CSetting::TYPE_BITMAP || pSetting->type==CSetting::TYPE_SOUND || pSetting->type==CSetting::TYPE_FONT)
		{
			if (pSetting->value.vt==VT_BSTR)
				Strcpy(text,_countof(text),pSetting->value.bstrVal);
			else
				text[0]=0;
			if (pSetting->type==CSetting::TYPE_STRING)
				mode=EDIT_STRING;
			else if (pSetting->type==CSetting::TYPE_ICON)
				mode=EDIT_ICON;
			else if (pSetting->type==CSetting::TYPE_BITMAP)
				mode=EDIT_BITMAP;
			else if (pSetting->type==CSetting::TYPE_SOUND)
				mode=EDIT_SOUND;
			else
				mode=EDIT_FONT;
		}
		else if (pSetting->type==CSetting::TYPE_HOTKEY || pSetting->type==CSetting::TYPE_HOTKEY_ANY)
		{
			text[0]=0;
			g_HotKey=0;
			if (pSetting->value.vt==VT_I4)
			{
				g_HotKey=pSetting->value.intVal;
				GetKeyName(pSetting->value.intVal,text,_countof(text),false,false,false);
			}
			if (bEnabled)
				mode=pSetting->type==CSetting::TYPE_HOTKEY_ANY?EDIT_HOTKEY_ANY:EDIT_HOTKEY;
		}
		else if (pSetting->type==CSetting::TYPE_COLOR)
		{
			mode=EDIT_COLOR;
			int val=0;
			if (pSetting->value.vt==VT_I4)
				val=pSetting->value.intVal;
			Sprintf(text,_countof(text),L"%06X",val);
		}
	}

	RECT rc;
	m_EditMode=mode;
	if (mode!=EDIT_NONE)
	{
		CString str=LoadStringEx(pSetting->nameID);
		TVITEM item={TVIF_TEXT,hItem,0,0,(LPWSTR)(LPCWSTR)str};
		TreeView_SetItem(m_Tree,&item);
		m_Tree.GetClientRect(&rc);
		int w=rc.right;
		TreeView_GetItemRect(m_Tree,hItem,&rc,TRUE);
		rc.left=rc.right;
		rc.right=w;
		m_EditItem=hItem;
		m_pEditSetting=pSetting;
	}

	if (mode==EDIT_ICON || mode==EDIT_BITMAP || mode==EDIT_SOUND || mode==EDIT_FONT || mode==EDIT_COLOR)
	{
		RECT rc2=rc;
		int h=rc2.bottom-rc2.top;
		rc2.left=rc2.right-h;
		if (mode==EDIT_SOUND)
		{
			m_PlayButton.SetWindowPos(NULL,&rc2,SWP_NOZORDER|SWP_SHOWWINDOW);
			rc2.left-=h;
			rc2.right-=h;
		}
		else
			m_PlayButton.ShowWindow(SW_HIDE);
		m_BrowseButton.EnableWindow(bEnabled);
		m_BrowseButton.SetWindowPos(NULL,&rc2,SWP_NOZORDER|SWP_SHOWWINDOW);
		rc.right=rc2.left;
	}
	else
	{
		m_BrowseButton.ShowWindow(SW_HIDE);
		m_PlayButton.ShowWindow(SW_HIDE);
	}

	if (mode!=EDIT_NONE)
	{
		m_EditBox.SetWindowText(text);
		m_EditBox.SendMessage(EM_SETREADONLY,!bEnabled,0);
		m_EditBox.SetWindowPos(NULL,&rc,SWP_NOZORDER|SWP_SHOWWINDOW);
		SendMessage(WM_NEXTDLGCTL,(LPARAM)m_EditBox.m_hWnd,TRUE);
		SetWindowSubclass(m_EditBox,SubclassEditProc,'CLSH',(mode==EDIT_HOTKEY)?1:((mode==EDIT_HOTKEY_ANY)?2:0));
	}
	else
		m_EditBox.ShowWindow(SW_HIDE);
}

void CTreeSettingsDlg::UpdateEditPosition( void )
{
	if (m_EditMode==EDIT_NONE) return;

	RECT rc;
	m_Tree.GetClientRect(&rc);
	int w=rc.right;
	TreeView_GetItemRect(m_Tree,m_EditItem,&rc,TRUE);
	rc.left=rc.right;
	rc.right=w;
	w-=rc.left;
	int h=rc.bottom-rc.top;

	HDC hdc=CreateCompatibleDC(NULL);
	HFONT font0=(HFONT)SelectObject(hdc,m_EditBox.GetFont());
	SIZE size;
	CString str;
	m_EditBox.GetWindowText(str);
	GetTextExtentPoint(hdc,str,str.GetLength(),&size);
	SelectObject(hdc,font0);
	DeleteDC(hdc);
	DWORD margins=(DWORD)m_EditBox.SendMessage(EM_GETMARGINS);
	size.cx+=HIWORD(margins)+LOWORD(margins)+12;
	if (m_EditMode==EDIT_ICON || m_EditMode==EDIT_BITMAP|| m_EditMode==EDIT_FONT || m_EditMode==EDIT_COLOR)
		size.cx+=h;
	if (m_EditMode==EDIT_SOUND)
		size.cx+=h*2;
	if (size.cx<w)
		rc.right=rc.left+size.cx;

	if (m_EditMode==EDIT_ICON || m_EditMode==EDIT_BITMAP || m_EditMode==EDIT_SOUND || m_EditMode==EDIT_FONT || m_EditMode==EDIT_COLOR)
	{
		RECT rc2=rc;
		rc2.left=rc2.right-h;
		if (m_EditMode==EDIT_SOUND)
		{
			m_PlayButton.SetWindowPos(NULL,&rc2,SWP_NOZORDER|SWP_NOCOPYBITS);
			rc2.left-=h;
			rc2.right-=h;
		}
		m_BrowseButton.SetWindowPos(NULL,&rc2,SWP_NOZORDER|SWP_NOCOPYBITS);
		rc.right=rc2.left;
	}

	m_EditBox.SetWindowPos(NULL,&rc,SWP_NOZORDER|SWP_NOCOPYBITS);
}

const CSetting *CTreeSettingsDlg::GetNextSetting( const CSetting *pSetting )
{
	pSetting++;
	if (pSetting->type==CSetting::TYPE_RADIO && !(pSetting->flags&CSetting::FLAG_HIDDEN))
		return pSetting;

	if (m_bBasic)
	{
		while (pSetting->name && (pSetting->type==CSetting::TYPE_GROUP || (pSetting->flags&CSetting::FLAG_HIDDEN) || !(pSetting->flags&CSetting::FLAG_BASIC)))
			pSetting++;
	}
	else
	{
		while (pSetting->name && pSetting->type!=CSetting::TYPE_GROUP && (pSetting->flags&CSetting::FLAG_HIDDEN))
			pSetting++;
	}

	if (pSetting->name && pSetting->type!=CSetting::TYPE_GROUP)
		return pSetting;
	return NULL;
}

void CTreeSettingsDlg::SetGroup( CSetting *pGroup )
{
	ItemSelected(NULL,NULL,false);
	m_pGroup=pGroup;
	// fill tree control
	m_Tree.SendMessage(WM_SETREDRAW,FALSE);
	TreeView_SelectItem(m_Tree,NULL);
	TreeView_DeleteAllItems(m_Tree);
	int level=0;
	HTREEITEM hParent=NULL;
	m_bBasic=(pGroup->nameID==IDS_BASIC_SETTINGS);
	for (const CSetting *pSetting=GetNextSetting(pGroup);pSetting;pSetting=GetNextSetting(pSetting))
	{
		wchar_t text[256];
		text[0]=0;
		if (pSetting->type==CSetting::TYPE_BOOL || pSetting->type==CSetting::TYPE_RADIO || (pSetting->type==CSetting::TYPE_INT && pSetting[1].type==CSetting::TYPE_RADIO))
		{
			// the text for BOOL, RADIO or RADIO's parent items doesn't change
			Strcpy(text,_countof(text),LoadStringEx(pSetting->nameID));
		}
		if (pSetting->type!=CSetting::TYPE_RADIO)
			hParent=NULL;
		TVINSERTSTRUCT insert={hParent,TVI_LAST,{TVIF_STATE|TVIF_TEXT|TVIF_IMAGE|TVIF_SELECTEDIMAGE|TVIF_PARAM,NULL,TVIS_EXPANDED,TVIS_EXPANDED,text,0,SETTING_STATE_SETTING,SETTING_STATE_SETTING}};
		insert.item.lParam=(LPARAM)pSetting;
		if (pSetting->flags&CSetting::FLAG_LOCKED_MASK)
		{
			insert.item.state|=INDEXTOOVERLAYMASK(1);
			insert.item.stateMask|=TVIS_OVERLAYMASK;
		}
		else if (pSetting->flags&CSetting::FLAG_WARNING)
		{
			insert.item.state|=INDEXTOOVERLAYMASK(2);
			insert.item.stateMask|=TVIS_OVERLAYMASK;
		}
		HTREEITEM hItem=TreeView_InsertItem(m_Tree,&insert);
		if (pSetting->type!=CSetting::TYPE_RADIO)
			hParent=hItem;
	}
	UpdateGroup(NULL);
	m_Tree.SendMessage(WM_SETREDRAW,TRUE);
	m_Tree.InvalidateRect(NULL);
}

void CTreeSettingsDlg::UpdateGroup( const CSetting *pModified )
{
	HTREEITEM hItem=NULL;
	for (const CSetting *pSetting=GetNextSetting(m_pGroup);pSetting;pSetting=GetNextSetting(pSetting))
	{
		if (pSetting->type==CSetting::TYPE_RADIO) continue;
		hItem=hItem?TreeView_GetNextSibling(m_Tree,hItem):TreeView_GetRoot(m_Tree);

		wchar_t text[256];
		TVITEM item={TVIF_STATE|TVIF_IMAGE,hItem,0,TVIS_BOLD|TVIS_CUT,text};
		TreeView_GetItem(m_Tree,&item);

		// check if the item is enabled
		bool bEnabled=pSetting->IsEnabled();

		// check if the item is default
		bool bDefault=(pSetting->flags&CSetting::FLAG_DEFAULT)!=0;

		// calculate text
		if (pSetting!=m_pEditSetting)
		{
			if (pSetting->type==CSetting::TYPE_HOTKEY || pSetting->type==CSetting::TYPE_HOTKEY_ANY)
			{
				wchar_t val[100];
				val[0]=0;
				if (pSetting->value.vt==VT_I4)
					GetKeyName(pSetting->value.intVal,val,_countof(val),false,false,false);
				CString str=LoadStringEx(pSetting->nameID);
				Sprintf(text,_countof(text),L"%s: %s",str,val);
				item.mask|=TVIF_TEXT;
			}
			else if (pSetting->type>=CSetting::TYPE_STRING)
			{
				CString str=LoadStringEx(pSetting->nameID);
				const wchar_t *val=L"";
				if (pSetting->value.vt==VT_BSTR)
					val=pSetting->value.bstrVal;
				Sprintf(text,_countof(text),L"%s: %s",str,val);
				item.mask|=TVIF_TEXT;
			}
			else if (pSetting->type==CSetting::TYPE_INT && pSetting[1].type!=CSetting::TYPE_RADIO)
			{
				CString str=LoadStringEx(pSetting->nameID);
				int val=0;
				if (pSetting->value.vt==VT_I4)
					val=pSetting->value.intVal;
				Sprintf(text,_countof(text),L"%s: %d",str,val);
				item.mask|=TVIF_TEXT;
			}
			else if (pSetting->type==CSetting::TYPE_COLOR)
			{
				CString str=LoadStringEx(pSetting->nameID);
				int val=0;
				if (pSetting->value.vt==VT_I4)
					val=pSetting->value.intVal;
				Sprintf(text,_countof(text),L"%s: %06X",str,val);
				item.mask|=TVIF_TEXT;
			}
		}

		// calculate state
		int image=SETTING_STATE_SETTING;
		if (pSetting->type==CSetting::TYPE_BOOL)
			image=SETTING_STATE_CHECKBOX|(IsVariantTrue(pSetting->value)?SETTING_STATE_CHECKED:0);
		if (pSetting->type==CSetting::TYPE_COLOR)
		{
			image=SETTING_IMAGE_COLOR+(pSetting->flags>>24);

			HIMAGELIST images=GetSettingsImageList(m_Tree);
			int cx, cy;
			ImageList_GetIconSize(images,&cx,&cy);

			BITMAPINFO dib={sizeof(dib)};
			dib.bmiHeader.biWidth=cx;
			dib.bmiHeader.biHeight=cy;
			dib.bmiHeader.biPlanes=1;
			dib.bmiHeader.biBitCount=32;
			dib.bmiHeader.biCompression=BI_RGB;
			HDC hdc=CreateCompatibleDC(NULL);
			HDC hdcMask=CreateCompatibleDC(NULL);
			HBITMAP bmp=CreateDIBSection(hdc,&dib,DIB_RGB_COLORS,NULL,NULL,0);
			HBITMAP bmpMask=CreateDIBSection(hdcMask,&dib,DIB_RGB_COLORS,NULL,NULL,0);

			HBITMAP bmp0=(HBITMAP)SelectObject(hdc,bmp);
			HBITMAP bmp1=(HBITMAP)SelectObject(hdcMask,bmpMask);
			SetDCBrushColor(hdc,pSetting->value.intVal&0xFFFFFF);
			SetDCPenColor(hdc,0);
			SelectObject(hdc,GetStockObject(DC_BRUSH));
			SelectObject(hdc,GetStockObject(DC_PEN));
			Rectangle(hdc,0,0,cx,cy);
			RECT rc={0,0,cx,cy};
			FillRect(hdcMask,&rc,(HBRUSH)GetStockObject(BLACK_BRUSH));
			SelectObject(hdc,bmp0);
			SelectObject(hdcMask,bmp1);
			ImageList_Replace(images,image,bmp,bmpMask);
			DeleteObject(bmp);
			DeleteObject(bmpMask);
			DeleteDC(hdc);
			DeleteDC(hdcMask);
		}
		int state=bDefault?0:TVIS_BOLD;
		if (!bEnabled)
		{
			if (pSetting->type!=CSetting::TYPE_COLOR) image|=SETTING_STATE_DISABLED;
			state|=TVIS_CUT;
		}
		if (item.iImage==image)
			item.mask&=~TVIF_IMAGE;
		else
			item.iImage=item.iSelectedImage=image;

		if ((item.state&(TVIS_BOLD|TVIS_CUT))==state)
			item.mask&=~TVIF_STATE;
		else
			item.state=state;
		if (item.mask!=0)
		{
			if (item.mask&TVIF_IMAGE)
				item.mask|=TVIF_SELECTEDIMAGE;
			TreeView_SetItem(m_Tree,&item);
			RECT rc;
			TreeView_GetItemRect(m_Tree,hItem,&rc,FALSE);
			m_Tree.InvalidateRect(&rc);
		}

		// update radio buttons
		if (pSetting->type==CSetting::TYPE_INT && pSetting[1].type==CSetting::TYPE_RADIO)
		{
			int val=0;
			if (pSetting->value.vt==VT_I4)
				val=pSetting->value.intVal;
			for (HTREEITEM hRadio=TreeView_GetChild(m_Tree,hItem);hRadio;hRadio=TreeView_GetNextSibling(m_Tree,hRadio))
			{
				TVITEM radio={TVIF_IMAGE|TVIF_PARAM|TVIF_STATE,hRadio,0,TVIS_CUT};
				TreeView_GetItem(m_Tree,&radio);
				int v=(int)((CSetting*)radio.lParam-pSetting-1);

				int image=SETTING_STATE_RADIO;
				if (v==val) image|=SETTING_STATE_CHECKED;
				int state=0;
				if (!bEnabled)
				{
					if (pSetting->type!=CSetting::TYPE_COLOR) image|=SETTING_STATE_DISABLED;
					state=TVIS_CUT;
				}

				if (radio.iImage!=image || radio.state!=state)
				{
					radio.iImage=radio.iSelectedImage=image;
					radio.mask=TVIF_IMAGE|TVIF_SELECTEDIMAGE|TVIF_STATE;
					radio.state=state;
					TreeView_SetItem(m_Tree,&radio);
					RECT rc;
					TreeView_GetItemRect(m_Tree,hRadio,&rc,FALSE);
					m_Tree.InvalidateRect(&rc);
				}
			}
		}
	}
}

class CDefaultSettingsPanel: public ISettingsPanel
{
public:
	virtual HWND Create( HWND parent );
	virtual HWND Activate( CSetting *pGroup, const RECT &rect, bool bReset );
	virtual bool Validate( HWND parent );

private:
	static CTreeSettingsDlg s_Dialog;
};

CTreeSettingsDlg CDefaultSettingsPanel::s_Dialog;

ISettingsPanel *GetDefaultSettings( void )
{
	static CDefaultSettingsPanel panel;
	return &panel;
}

HWND CDefaultSettingsPanel::Create( HWND parent )
{
	if (!s_Dialog.m_hWnd)
		s_Dialog.Create(parent,LoadDialogEx(IDD_SETTINGSTREE));
	return s_Dialog.m_hWnd;
}

HWND CDefaultSettingsPanel::Activate( CSetting *pGroup, const RECT &rect, bool bReset )
{
	s_Dialog.SetGroup(pGroup);
	s_Dialog.SetWindowPos(HWND_TOP,&rect,SWP_SHOWWINDOW);
	return s_Dialog.m_hWnd;
}

bool CDefaultSettingsPanel::Validate( HWND parent )
{
	s_Dialog.Validate();
	return true;
}

///////////////////////////////////////////////////////////////////////////////

class CLanguageSettingsDlg: public CResizeableDlg<CLanguageSettingsDlg>
{
public:
	BEGIN_MSG_MAP( CLanguageSettingsDlg )
		MESSAGE_HANDLER( WM_INITDIALOG, OnInitDialog )
		MESSAGE_HANDLER( WM_DESTROY, OnDestroy )
		MESSAGE_HANDLER( WM_SIZE, OnSize )
		COMMAND_HANDLER( IDC_LISTLANGUAGE, LBN_SELCHANGE, OnSelChange )
	END_MSG_MAP()

	BEGIN_RESIZE_MAP
		RESIZE_CONTROL(IDC_LISTLANGUAGE,MOVE_SIZE_X|MOVE_SIZE_Y)
	END_RESIZE_MAP

	void SetGroup( CSetting *pGroup );

protected:
	// Handler prototypes:
	//  LRESULT MessageHandler(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
	//  LRESULT CommandHandler(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled);
	//  LRESULT NotifyHandler(int idCtrl, LPNMHDR pnmh, BOOL& bHandled);
	LRESULT OnInitDialog( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled );
	LRESULT OnDestroy( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled );
	LRESULT OnSize( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled );
	LRESULT OnSelChange( WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled );

private:
	CSetting *m_pSetting;
	CWindow m_Tooltip;
};

const int g_LanguageIDs[]={
	0x0401,
	0x0402,
	0x0405,
	0x0406,
	0x0407,
	0x0408,
	0x0409,
	0x0C0A,
	0x0425,
	0x040B,
	0x040C,
	0x040D,
	0x041A,
	0x040E,
	0x0410,
	0x0411,
	0x0412,
	0x0427,
	0x0426,
	0x0414,
	0x0413,
	0x0415,
	0x0416,
	0x0816,
	0x0418,
	0x0419,
	0x041B,
	0x0424,
	0x081A,
	0x041D,
	0x041E,
	0x041F,
	0x0422,
	0x0804,
	0x0C04,
	0x0404,
};

LRESULT CLanguageSettingsDlg::OnInitDialog( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled )
{
	CResizeableDlg<CLanguageSettingsDlg>::InitResize();
	EnableThemeDialogTexture(m_hWnd,ETDT_ENABLETAB);

	CWindow list=GetDlgItem(IDC_LISTLANGUAGE);
	list.SendMessage(LB_ADDSTRING,0,(LPARAM)(const wchar_t*)LoadStringEx(IDS_OS_LANGUAGE));

	std::vector<CString> names;
	for (int i=0;i<_countof(g_LanguageIDs);i++)
	{
		int id=g_LanguageIDs[i];
		wchar_t name1[256];
		wchar_t name2[256];
		GetLocaleInfo(id,LOCALE_SNAME,name1,_countof(name1));
		GetLocaleInfo(id,LOCALE_SLANGUAGE,name2,_countof(name2));
		wchar_t name[256];
		Sprintf(name,_countof(name),L"%s\xA0- %s",name1,name2);
		names.push_back(name);
	}
	std::stable_sort(names.begin(),names.end());
	for (std::vector<CString>::const_iterator it=names.begin();it!=names.end();++it)
		list.SendMessage(LB_ADDSTRING,0,(LPARAM)(const wchar_t*)*it);

	m_Tooltip.Create(TOOLTIPS_CLASS,m_hWnd,NULL,NULL,WS_POPUP|TTS_NOPREFIX);
	TOOLINFO tool={sizeof(tool),TTF_SUBCLASS,m_hWnd,'CLSH'};
	CString str=LoadStringEx(IDS_SETTING_LOCKED);
	tool.lpszText=(LPWSTR)(LPCWSTR)str;
	list.GetClientRect(&tool.rect);
	list.ClientToScreen(&tool.rect);
	ScreenToClient(&tool.rect);
	m_Tooltip.SendMessage(TTM_ADDTOOL,0,(LPARAM)&tool);

	return TRUE;
}

LRESULT CLanguageSettingsDlg::OnDestroy( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled )
{
	m_Tooltip.Detach();
	return 0;
}

LRESULT CLanguageSettingsDlg::OnSize( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled )
{
	CResizeableDlg<CLanguageSettingsDlg>::OnSize();
	CWindow list=GetDlgItem(IDC_LISTLANGUAGE);
	TOOLINFO tool={sizeof(tool),TTF_SUBCLASS,m_hWnd,'CLSH'};
	list.GetClientRect(&tool.rect);
	list.ClientToScreen(&tool.rect);
	ScreenToClient(&tool.rect);
	m_Tooltip.SendMessage(TTM_NEWTOOLRECT,0,(LPARAM)&tool);
	return 0;
}

LRESULT CLanguageSettingsDlg::OnSelChange( WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled )
{
	// set setting
	wchar_t name[256];
	CWindow list=GetDlgItem(IDC_LISTLANGUAGE);
	int idx=(int)list.SendMessage(LB_GETCURSEL);
	if (idx<0) idx=0;
	list.SendMessage(LB_GETTEXT,idx,(LPARAM)name);
	wchar_t *end=wcschr(name,(unsigned char)'\xA0');
	CSettingsLockWrite lock;
	if (end)
	{
		*end=0;
		m_pSetting->value=CComVariant(name);
	}
	else
		m_pSetting->value=CComVariant(L"");

	if (_wcsicmp(m_pSetting->value.bstrVal,m_pSetting->defValue.bstrVal)==0)
		m_pSetting->flags|=CSetting::FLAG_DEFAULT;
	else
		m_pSetting->flags&=~CSetting::FLAG_DEFAULT;
	return 0;
}

void CLanguageSettingsDlg::SetGroup( CSetting *pGroup )
{
	m_pSetting=pGroup+1;

	const wchar_t *language=L"";
	if (m_pSetting->value.vt==VT_BSTR && m_pSetting->value.bstrVal)
		language=m_pSetting->value.bstrVal;
	int len=Strlen(language);

	CWindow list=GetDlgItem(IDC_LISTLANGUAGE);
	int n=(int)list.SendMessage(LB_GETCOUNT);
	for (int i=0;i<n;i++)
	{
		wchar_t name[256];
		list.SendMessage(LB_GETTEXT,i,(LPARAM)name);
		wchar_t *end=wcschr(name,(unsigned char)'\xA0');
		if (end)
			*end=0;
		else
			*name=0;
		if (_wcsicmp(language,name)==0)
		{
			list.SendMessage(LB_SETCURSEL,i);
			break;
		}
	}
	list.EnableWindow(!(m_pSetting->flags&CSetting::FLAG_LOCKED_MASK));
}

class CLanguageSettingsPanel: public ISettingsPanel
{
public:
	virtual HWND Create( HWND parent );
	virtual HWND Activate( CSetting *pGroup, const RECT &rect, bool bReset );
	virtual bool Validate( HWND parent ) { return true; }

private:
	static CLanguageSettingsDlg s_Dialog;
};

CLanguageSettingsDlg CLanguageSettingsPanel::s_Dialog;

ISettingsPanel *GetLanguageSettings( void )
{
	static CLanguageSettingsPanel panel;
	return &panel;
}

HWND CLanguageSettingsPanel::Create( HWND parent )
{
	if (!s_Dialog.m_hWnd)
		s_Dialog.Create(parent,LoadDialogEx(IDD_LANGUAGE));
	return s_Dialog.m_hWnd;
}

HWND CLanguageSettingsPanel::Activate( CSetting *pGroup, const RECT &rect, bool bReset )
{
	s_Dialog.SetGroup(pGroup);
	s_Dialog.SetWindowPos(HWND_TOP,&rect,SWP_SHOWWINDOW);
	return s_Dialog.m_hWnd;
}

///////////////////////////////////////////////////////////////////////////////

static tNewVersionCallback g_NewVersionCallback;
static bool g_bCheckingVersion;

static DWORD WINAPI ThreadVersionCheck( void *param )
{
	ULONGLONG curTimeL;
	GetSystemTimeAsFileTime((FILETIME*)&curTimeL);
	DWORD curTime=(DWORD)(curTimeL/36000000000); // in hours

	CRegKey regKey;
	if (regKey.Open(HKEY_CURRENT_USER,L"Software\\IvoSoft\\ClassicShell")!=ERROR_SUCCESS)
	{
		g_bCheckingVersion=false;
		return 0;
	}
	DWORD newVersion;
	CString downloadUrl, news;
	TVersionCheck check=(TVersionCheck)(int)param;
	CString url=LoadStringEx(IDS_VERSION_URL);
	if (GetWinVersion()==WIN_VER_VISTA)
		url+=".V";
	{
		CRegKey regKeyLng;
		wchar_t language[100]=L".";
		if (regKeyLng.Open(HKEY_LOCAL_MACHINE,L"Software\\IvoSoft\\ClassicShell",KEY_READ|KEY_WOW64_64KEY)==ERROR_SUCCESS)
		{
			ULONG size=_countof(language)-1;
			if (regKeyLng.QueryStringValue(L"DefaultLanguage",language+1,&size)==ERROR_SUCCESS && size>1)
				url+=language;
		}
	}
	bool res=false;
	HINTERNET hInternet=InternetOpen(L"Classic Shell",INTERNET_OPEN_TYPE_PRECONFIG,NULL,NULL,0);
	if (hInternet)
	{
		wchar_t header[]=L"Accept: */*\r\n\r\n";
		HINTERNET hUrl=InternetOpenUrl(hInternet,url,header,Strlen(header),INTERNET_FLAG_RELOAD,0);
		if (hUrl)
		{
			DWORD dwSize;
			wchar_t buffer[4097];
			if (InternetReadFile(hUrl,buffer,sizeof(buffer)-2,&dwSize))
			{
				buffer[dwSize/2]=0;
				wchar_t token[256];
				const wchar_t *str=buffer;
				if (*str==0xFEFF) str++;
				str=GetToken(str,token,_countof(token),L"\r\n");
				int v1, v2, v3;
				if (swscanf_s(token,L"%d.%d.%d",&v1,&v2,&v3)==3)
				{
					newVersion=(v1<<24)|(v2<<16)|v3;
					str=GetToken(str,token,_countof(token),L"\r\n");
					downloadUrl=token;
					str=wcsstr(str,L"{NEWS}");
					news=str?str+6:L"";
					res=true;
				}
			}
			InternetCloseHandle(hUrl);
		}
		InternetCloseHandle(hInternet);
	}
	if (!res)
	{
		g_bCheckingVersion=false;
		return 0;
	}

	regKey.SetDWORDValue(L"LastUpdateTime",curTime);
	if (check==CHECK_UPDATE)
	{
		g_NewVersionCallback(newVersion,downloadUrl,news);
		g_bCheckingVersion=false;
		return 1;
	}

	DWORD remindedVersion;
	if (regKey.QueryDWORDValue(L"RemindedVersion",remindedVersion)!=ERROR_SUCCESS)
		remindedVersion=0;
	if (newVersion>remindedVersion)
		g_NewVersionCallback(newVersion,downloadUrl,news);
	g_bCheckingVersion=false;
	return 0;
}

bool CheckForNewVersion( TVersionCheck check, tNewVersionCallback callback )
{
	if (g_bCheckingVersion) return false;
	if (check!=CHECK_UPDATE)
	{
		// check admin settings
		CRegKey regKey;
		if (regKey.Open(HKEY_LOCAL_MACHINE,L"Software\\IvoSoft\\ClassicShell",KEY_READ|KEY_WOW64_64KEY)==ERROR_SUCCESS)
		{
			DWORD update;
			if (regKey.QueryDWORDValue(L"Update",update)==ERROR_SUCCESS && update==0)
				return false;
		}
	}

	wchar_t path[_MAX_PATH];
	GetModuleFileName(_AtlBaseModule.GetModuleInstance(),path,_countof(path));
	PathRemoveFileSpec(path);
	PathAppend(path,L"ClassicShellUpdate.exe");
	if (GetFileAttributes(path)==INVALID_FILE_ATTRIBUTES)
		return false;

	CRegKey regKey;
	if (regKey.Open(HKEY_CURRENT_USER,L"Software\\IvoSoft\\ClassicShell")!=ERROR_SUCCESS)
		regKey.Create(HKEY_CURRENT_USER,L"Software\\IvoSoft\\ClassicShell");

	ULONGLONG curTimeL;
	GetSystemTimeAsFileTime((FILETIME*)&curTimeL);
	DWORD curTime=(DWORD)(curTimeL/36000000000); // in hours

	g_NewVersionCallback=callback;
	if (check!=CHECK_UPDATE)
	{
		DWORD update;
		if (regKey.QueryDWORDValue(L"Update",update)==ERROR_SUCCESS && update==0)
			return false;

		DWORD lastTime;
		if (regKey.QueryDWORDValue(L"LastUpdateTime",lastTime)!=ERROR_SUCCESS)
			lastTime=0;
		if ((curTime-lastTime)<168)
			return false; // check weekly

		g_bCheckingVersion=true;
		if (check==CHECK_AUTO_WAIT)
			return ThreadVersionCheck((void*)check)!=0;
		HANDLE hThread=CreateThread(NULL,0,ThreadVersionCheck,(void*)check,0,NULL);
		CloseHandle(hThread);
		return hThread!=NULL;
	}

	g_bCheckingVersion=true;
	return ThreadVersionCheck((void*)check)!=0;
}
