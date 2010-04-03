// Classic Shell (c) 2009-2010, Ivo Beltchev
// The sources for Classic Shell are distributed under the MIT open source license

#pragma once

#include "SkinManager.h"
#include <vector>
#include <map>

#define ALLOW_DEACTIVATE // undefine this to prevent the menu from closing when it is deactivated (useful for debugging)
//#define REPEAT_ITEMS 10 // define this to repeat each menu item (useful to simulate large menus)

#if defined(BUILD_SETUP) && !defined(ALLOW_DEACTIVATE)
#define ALLOW_DEACTIVATE // make sure it is defined in Setup
#endif

#if defined(BUILD_SETUP) && defined(REPEAT_ITEMS)
#undef REPEAT_ITEMS
#endif

enum TMenuID
{
	MENU_NO=0,
	MENU_LAST=0,
	MENU_SEPARATOR,
	MENU_EMPTY,
	MENU_COLUMN_PADDING,
	MENU_COLUMN_BREAK,

	// standard menu items
	MENU_PROGRAMS,
	MENU_FAVORITES,
	MENU_DOCUMENTS,
		MENU_USERFILES,
		MENU_USERDOCUMENTS,
		MENU_USERPICTURES,
	MENU_SETTINGS,
		MENU_CONTROLPANEL,
		MENU_NETWORK,
		MENU_PRINTERS,
		MENU_TASKBAR,
		MENU_FEATURES,
		MENU_CLASSIC_SETTINGS,
	MENU_SEARCH,
		MENU_SEARCH_FILES,
		MENU_SEARCH_PRINTER,
		MENU_SEARCH_COMPUTERS,
		MENU_SEARCH_PEOPLE,
	MENU_HELP,
	MENU_RUN,
	MENU_LOGOFF,
	MENU_DISCONNECT,
	MENU_UNDOCK,
	MENU_SHUTDOWN_BOX,

	// additional commands
	MENU_CUSTOM, // used for any custom item
	MENU_SLEEP,
	MENU_HIBERNATE,
	MENU_RESTART,
	MENU_SHUTDOWN,
	MENU_SWITCHUSER,
};

struct StdMenuItem
{
	TMenuID id;
	const char *key; // localization key
	const wchar_t *name; // default name
	int icon; // index in shell32.dll
	TMenuID submenuID; // MENU_NO if no submenu
	const KNOWNFOLDERID *folder1; // NULL if not used
	const KNOWNFOLDERID *folder2; // NULL if not used
	const char *tipKey; // localization key for the tooltip
	const wchar_t *tip; // default tooltip
	const StdMenuItem *submenu;
	const wchar_t *link;
	const wchar_t *command;
	const wchar_t *iconPath;

	// user settings
	enum
	{
		MENU_OPENUP     = 0x0001, // prefer to open up
		MENU_OPENUP_REC = 0x0002, // children prefer to open up
		MENU_SORTZA     = 0x0004, // sort backwards
		MENU_SORTZA_REC = 0x0008, // children sort backwards
		MENU_SORTONCE   = 0x0010, // save the sort order the first time the menu is opened
	};
	unsigned int settings;
};

class CMenuAccessible;

// CMenuContainer - implementation of a single menu box.
class CMenuContainer: public CWindowImpl<CMenuContainer>, public IDropTarget
{
public:
	DECLARE_WND_CLASS_EX(L"ClassicShell.CMenuContainer",CS_DROPSHADOW|CS_DBLCLKS,COLOR_MENU)

	// message handlers
	BEGIN_MSG_MAP( CMenuContainer )
		// forward all messages to m_pMenu2 and m_pMenu3 to ensure the context menu functions properly
		if (m_pMenu3)
		{
			if (SUCCEEDED(m_pMenu3->HandleMenuMsg2(uMsg,wParam,lParam,&lResult)))
				return TRUE;
		}
		else if (m_pMenu2)
		{
			if (SUCCEEDED(m_pMenu2->HandleMenuMsg(uMsg,wParam,lParam)))
			{
				lResult=0;
				return TRUE;
			}
		}
		MESSAGE_HANDLER( WM_CREATE, OnCreate )
		MESSAGE_HANDLER( WM_DESTROY, OnDestroy )
		MESSAGE_HANDLER( WM_PAINT, OnPaint )
		MESSAGE_HANDLER( WM_PRINTCLIENT, OnPaint )
		MESSAGE_HANDLER( WM_ERASEBKGND, OnEraseBkgnd )
		MESSAGE_HANDLER( WM_ACTIVATE, OnActivate )
		MESSAGE_HANDLER( WM_MOUSEACTIVATE, OnMouseActivate )
		MESSAGE_HANDLER( WM_MOUSEMOVE, OnMouseMove )
		MESSAGE_HANDLER( WM_MOUSELEAVE, OnMouseLeave )
		MESSAGE_HANDLER( WM_MOUSEWHEEL, OnMouseWheel )
		MESSAGE_HANDLER( WM_LBUTTONDOWN, OnLButtonDown )
		MESSAGE_HANDLER( WM_LBUTTONDBLCLK, OnLButtonDblClick )
		MESSAGE_HANDLER( WM_LBUTTONUP, OnLButtonUp )
		MESSAGE_HANDLER( WM_RBUTTONDOWN, OnRButtonDown )
		MESSAGE_HANDLER( WM_RBUTTONUP, OnRButtonUp )
		MESSAGE_HANDLER( WM_SETCURSOR, OnSetCursor )
		MESSAGE_HANDLER( WM_CONTEXTMENU, OnContextMenu )
		MESSAGE_HANDLER( WM_KEYDOWN, OnKeyDown )
		MESSAGE_HANDLER( WM_CHAR, OnChar )
		MESSAGE_HANDLER( WM_TIMER, OnTimer )
		MESSAGE_HANDLER( WM_SYSCOMMAND, OnSysCommand )
		MESSAGE_HANDLER( WM_GETOBJECT, OnGetAccObject )
		MESSAGE_HANDLER( MCM_REFRESH, OnRefresh )
		MESSAGE_HANDLER( MCM_SETCONTEXTITEM, OnSetContextItem )
	END_MSG_MAP()

	// options when creating a container
	enum
	{
		CONTAINER_LARGE        = 0x0001, // use large icons
		CONTAINER_MULTICOLUMN  = 0x0002, // use multiple columns instead of a single scrolling column
		CONTAINER_CONTROLPANEL = 0x0004, // this is the control panel, don't go into subfolders
		CONTAINER_PROGRAMS     = 0x0008, // this is a folder from the Start Menu hierarchy (drop operations prefer link over move)
		CONTAINER_DOCUMENTS    = 0x0010, // sort by time, limit the count (for recent documents)
		CONTAINER_LINK         = 0x0020, // this is an expanded link to a folder (always scrolling)
		CONTAINER_ADDTOP       = 0x0040, // put standard items at the top
		CONTAINER_DRAG         = 0x0080, // allow items to be dragged out
		CONTAINER_DROP         = 0x0100, // allow dropping of items
		CONTAINER_LEFT         = 0x0200, // the window is aligned on the left
		CONTAINER_TOP          = 0x0400, // the window is aligned on the top
		CONTAINER_AUTOSORT     = 0x0800, // the menu is always in alphabetical order
		CONTAINER_OPENUP_REC   = 0x1000, // the container's children will prefer to open up instead of down
		CONTAINER_SORTZA       = 0x2000, // the container will sort backwards by default
		CONTAINER_SORTZA_REC   = 0x4000, // the container's children will sort backwards by default
		CONTAINER_SORTONCE     = 0x8000, // the container will save the sort order the first time the menu is opened
	};

	CMenuContainer( CMenuContainer *pParent, int index, int options, const StdMenuItem *pStdItem, PIDLIST_ABSOLUTE path1, PIDLIST_ABSOLUTE path2, const CString &regName );
	~CMenuContainer( void );

	void InitItems( void );
	void InitWindow( void );

	static bool CloseStartMenu( void );
	static void HideStartMenu( void );
	static bool IsMenuOpened( void ) { return !s_Menus.empty(); }
	static bool IgnoreTaskbarTimers( void ) { return !s_Menus.empty() && (s_TaskbarState&ABS_AUTOHIDE); }
	static HWND ToggleStartMenu( HWND startButton, bool bKeyboard );

	// IUnknown
	virtual STDMETHODIMP QueryInterface( REFIID riid, void **ppvObject )
	{
		*ppvObject=NULL;
		if (IID_IUnknown==riid || IID_IDropTarget==riid)
		{
			AddRef();
			*ppvObject=static_cast<IDropTarget*>(this);
			return S_OK;
		}
		return E_NOINTERFACE;
	}

	virtual ULONG STDMETHODCALLTYPE AddRef( void ) 
	{ 
		return InterlockedIncrement(&m_RefCount);
	}

	virtual ULONG STDMETHODCALLTYPE Release( void )
	{
		long nTemp=InterlockedDecrement(&m_RefCount);
		if (!nTemp) delete this;
		return nTemp;
	}

	// IDropTarget
	virtual HRESULT STDMETHODCALLTYPE DragEnter( IDataObject *pDataObj, DWORD grfKeyState, POINTL pt, DWORD *pdwEffect );
	virtual HRESULT STDMETHODCALLTYPE DragOver( DWORD grfKeyState, POINTL pt, DWORD *pdwEffect );
	virtual HRESULT STDMETHODCALLTYPE DragLeave( void );
	virtual HRESULT STDMETHODCALLTYPE Drop( IDataObject *pDataObj, DWORD grfKeyState, POINTL pt, DWORD *pdwEffect );

protected:
	LRESULT OnCreate( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled );
	LRESULT OnDestroy( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled );
	LRESULT OnRefresh( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled );
	LRESULT OnPaint( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled );
	LRESULT OnEraseBkgnd( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled ) { return 1; }
	LRESULT OnActivate( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled );
	LRESULT OnMouseActivate( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled );
	LRESULT OnMouseMove( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled );
	LRESULT OnMouseLeave( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled );
	LRESULT OnMouseWheel( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled );
	LRESULT OnLButtonDown( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled );
	LRESULT OnLButtonDblClick( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled );
	LRESULT OnLButtonUp( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled );
	LRESULT OnRButtonDown( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled );
	LRESULT OnRButtonUp( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled );
	LRESULT OnSetCursor( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled );
	LRESULT OnContextMenu( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled );
	LRESULT OnKeyDown( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled );
	LRESULT OnChar( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled );
	LRESULT OnTimer( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled );
	LRESULT OnSysCommand( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled );
	LRESULT OnGetAccObject( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled );
	LRESULT OnSetContextItem( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled );
	virtual void OnFinalMessage( HWND ) { Release(); }

private:
	// description of a menu item
	struct MenuItem
	{
		TMenuID id; // if pStdItem!=NULL, this is pStdItem->id. otherwise it can only be MENU_NO, MENU_SEPARATOR or MENU_EMPTY
		const StdMenuItem *pStdItem; // NULL if not a standard menu item
		CString name;
		unsigned int nameHash;
		int icon;
		int column;
		int row;
		RECT itemRect;
		bool bFolder:1; // this is a folder - draw arrow
		bool bLink:1; // this is a link (if a link to a folder is expanded it is always single-column)
		bool bPrograms:1; // this item is part of the Start Menu folder hierarchy
		bool bAlignBottom:1; // two-column menu: this item is aligned to the bottom
		bool bBreak:1; // two-column menu: this item starts the second column

		// pair of shell items. 2 items are used to combine a user folder with a common folder (I.E. user programs/common programs)
		PIDLIST_ABSOLUTE pItem1;
		PIDLIST_ABSOLUTE pItem2;
		union
		{
			UINT accelerator; // accelerator character, 0 if none
			FILETIME time; // timestamp of the file (for sorting recent documents)
		};

		bool operator<( const MenuItem &x ) const
		{
			if (row<x.row) return true;
			if (row>x.row) return false;
			if (bFolder && !x.bFolder) return true;
			if (!bFolder && x.bFolder) return false;
			if (bFolder)
			{
				const wchar_t *drive1=name.IsEmpty()?NULL:wcschr((const wchar_t*)name+1,':');
				const wchar_t *drive2=x.name.IsEmpty()?NULL:wcschr((const wchar_t*)x.name+1,':');
				if (drive1 && !drive2) return true;
				if (!drive1 && drive2) return false;
				if (drive1)
					return drive1[-1]<drive2[-1];
			}
			return CompareString(LOCALE_USER_DEFAULT,LINGUISTIC_IGNORECASE,name,-1,x.name,-1)==CSTR_LESS_THAN;
		}
	};

	struct SortMenuItem
	{
		CString name;
		unsigned int nameHash;
		bool bFolder;

		bool operator<( const SortMenuItem &x ) const
		{
			if (bFolder && !x.bFolder) return true;
			if (!bFolder && x.bFolder) return false;
			if (bFolder)
			{
				const wchar_t *drive1=name.IsEmpty()?NULL:wcschr((const wchar_t*)name+1,':');
				const wchar_t *drive2=x.name.IsEmpty()?NULL:wcschr((const wchar_t*)x.name+1,':');
				if (drive1 && !drive2) return true;
				if (!drive1 && drive2) return false;
				if (drive1)
					return drive1[-1]<drive2[-1];
			}
			return CompareString(LOCALE_USER_DEFAULT,LINGUISTIC_IGNORECASE,name,-1,x.name,-1)==CSTR_LESS_THAN;
		}
	};

	LONG m_RefCount;
	bool m_bSubMenu;
	bool m_bRefreshPosted;
	bool m_bDestroyed; // the menu is destroyed but not yet deleted
	bool m_bTrackMouse;
	int m_Options;
	const StdMenuItem *m_pStdItem; // the first item
	CMenuContainer *m_pParent; // parent menu
	int m_ParentIndex; // the index of this menu in the parent (usually matches m_pParent->m_Submenu)
	int m_Submenu; // the item index of the opened submenu
	int m_HotItem;
	int m_InsertMark;
	bool m_bInsertAfter;
	CString m_RegName; // name of the registry key to store the item order
	PIDLIST_ABSOLUTE m_Path1;
	PIDLIST_ABSOLUTE m_Path2;
	CComPtr<IShellFolder> m_pDropFolder; // the primary folder (used only as a drop target)
	CMenuAccessible *m_pAccessible;
	std::vector<int> m_ColumnOffsets;

	std::vector<MenuItem> m_Items; // all items in the menu (including separators)
	CComQIPtr<IContextMenu2> m_pMenu2; // additional interfaces used when a context menu is displayed
	CComQIPtr<IContextMenu3> m_pMenu3;

	int m_DragHoverTime;
	int m_DragHoverItem;
	int m_DragIndex; // the index of the item being dragged
	CComPtr<IDropTargetHelper> m_pDropTargetHelper; // to show images while dragging
	CComPtr<IDataObject> m_pDragObject;

	int m_ClickIndex; // the index of the last clicked item
	DWORD m_HotPos; // last mouse position over a hot item (used to ignore WM_MOUSEMOVE when the mouse didn't really move)
	int m_HoverItem; // item under the mouse (used for opening a submenu when the mouse hovers over an item)
	int m_ContextItem; // force this to be the hot item while a context menu is up
	HBITMAP m_Bitmap; // the background bitmap
	HBITMAP m_ArrowsBitmap[4]; // normal, selected, normal2, selected2
	HFONT m_Font[2];
	HRGN m_Region; // the outline region
	int m_MaxWidth;
	bool m_bTwoColumns;
	RECT m_rContent;
	RECT m_rContent2;
	int m_ItemHeight[2];
	int m_MaxItemWidth[2];
	int m_IconTopOffset[2]; // offset from the top of the item to the top of the icon
	int m_TextTopOffset[2]; // offset from the top of the item to the top of the text
	RECT m_rUser; // the user image (0,0,0,0 if the user image is not shown)

	int m_ScrollCount; // number of items to scroll in the pager
	int m_ScrollHeight; // 0 - don't scroll
	int m_ScrollOffset;
	int m_ScrollButtonSize;
	int m_MouseWheel;
	bool m_bScrollUp, m_bScrollDown;
	bool m_bScrollUpHot, m_bScrollDownHot;
	bool m_bScrollTimer;

	// additional commands for the context menu
	enum
	{
		CMD_OPEN_ALL=1,
		CMD_SORT,
		CMD_AUTOSORT,
		CMD_NEWFOLDER,

		CMD_LAST
	};

	// ways to activate a menu item
	enum TActivateType
	{
		ACTIVATE_SELECT, // just selects the item
		ACTIVATE_OPEN, // opens the submenu or selects if not a menu
		ACTIVATE_OPEN_KBD, // same as above, but when done with a keyboard
		ACTIVATE_EXECUTE, // executes the item
		ACTIVATE_MENU, // shows context menu
	};

	// sound events
	enum TMenuSound
	{
		SOUND_MAIN,
		SOUND_POPUP,
		SOUND_COMMAND,
		SOUND_DROP,
	};

	enum
	{
		// timer ID
		TIMER_HOVER=1,
		TIMER_SCROLL=2,
		TIMER_TOOLTIP_SHOW=3,
		TIMER_TOOLTIP_HIDE=4,
		TIMER_BALLOON_HIDE=5,

		MCM_REFRESH=WM_USER+10, // posted to force the container to refresh its contents
		MCM_SETCONTEXTITEM=WM_USER+11, // sets the item for the context menu. wParam is the nameHash of the item

		// some constants
		SEPARATOR_HEIGHT=8,
		MIN_SCROLL_HEIGHT=13, // the scroll buttons are at least this tall
		MAX_MENU_ITEMS=500,
		MENU_ANIM_SPEED=200,
		MENU_ANIM_SPEED_SUBMENU=100,
		MENU_FADE_SPEED=400,
		USER_PICTURE_SIZE=48,
	};

	// pPt - optional point in screen space (used only by ACTIVATE_EXECUTE and ACTIVATE_MENU)
	void ActivateItem( int index, TActivateType type, const POINT *pPt );
	void RunUserCommand( void );
	void ShowKeyboardCues( void );
	void SetActiveWindow( void );
	void CreateBackground( int width1, int width2, int height1, int height2 ); // width1/2, height1/2 - the first and second content area
	void CreateSubmenuRegion( int width, int height ); // width, height - the content area
	void PostRefreshMessage( void );
	void SaveItemOrder( const std::vector<SortMenuItem> &items );
	void LoadItemOrder( void );
	void FadeOutItem( int index );
	bool GetItemRect( int index, RECT &rc );
	int HitTest( const POINT &pt );
	bool DragOut( int index );
	void InvalidateItem( int index );
	void SetHotItem( int index );
	void SetInsertMark( int index, bool bAfter );
	bool GetInsertRect( RECT &rc );
	void DrawBackground( HDC hdc, const RECT &drawRect );
	bool GetDescription( int index, wchar_t *text, int size );
	void UpdateScroll( void );
	void UpdateScroll( const POINT *pt );
	void PlayMenuSound( TMenuSound sound );

	static int s_MaxRecentDocuments; // limit for the number of recent documents
	static int s_ScrollMenus; // global scroll menus setting
	static bool s_bRTL; // RTL layout
	static bool s_bKeyboardCues; // show keyboard cues
	static bool s_bExpandRight; // prefer expanding submenus to the right
	static bool s_bBehindTaskbar; // the main menu is behind the taskbar (when the taskbar is horizontal)
	static bool s_bShowTopEmpty; // shows the empty item on the top menu so the user can drag items there
	static bool s_bNoDragDrop; // disables drag/drop
	static bool s_bNoContextMenu; // disables the context menu
	static bool s_bExpandLinks; // expand links to folders
	static char s_bActiveDirectory; // the Active Directory services are available (-1 - uninitialized)
	static CMenuContainer *s_pDragSource; // the source of the current drag operation
	static bool s_bRightDrag; // dragging with the right mouse button
	static RECT s_MainRect; // area of the main monitor
	static DWORD s_TaskbarState; // the state of the taskbar (ABS_AUTOHIDE and ABS_ALWAYSONTOP)
	static DWORD s_HoverTime;
	static DWORD s_SubmenuStyle;
	static CLIPFORMAT s_ShellFormat; // CFSTR_SHELLIDLIST
	static CComPtr<IShellFolder> s_pDesktop; // cached pointer of the desktop object
	static HWND s_LastFGWindow; // stores the foreground window to restore later when the menu closes
	static HTHEME s_Theme;
	static HTHEME s_PagerTheme;
	static CWindow s_Tooltip;
	static CWindow s_TooltipBaloon;
	static int s_TipShowTime;
	static int s_TipHideTime;
	static int s_TipShowTimeFolder;
	static int s_TipHideTimeFolder;
	static int s_HotItem;
	static CMenuContainer *s_pHotMenu; // the menu with the hot item
	static int s_TipItem; // the item that needs a tooltip
	static CMenuContainer *s_pTipMenu;

	static std::vector<CMenuContainer*> s_Menus; // all menus, in cascading order
	static std::map<unsigned int,int> s_MenuScrolls; // scroll offset for each sub menu

	static MenuSkin s_Skin;

	friend class COwnerWindow;
	friend class CMenuAccessible;

	static void MarginsBlit( HDC hSrc, HDC hDst, const RECT &rSrc, const RECT &rDst, const RECT &rMargins, bool bAlpha, bool bRtlOffset=false );
};

class CMenuFader: public CWindowImpl<CMenuFader>
{
public:
	CMenuFader( HBITMAP bmp, HRGN region, int duration, RECT &rect );
	~CMenuFader( void );
	DECLARE_WND_CLASS_EX(L"ClassicShell.CMenuFader",0,COLOR_MENU)

	// message handlers
	BEGIN_MSG_MAP( CMenuFader )
		MESSAGE_HANDLER( WM_ERASEBKGND, OnEraseBkgnd )
		MESSAGE_HANDLER( WM_TIMER, OnTimer )
	END_MSG_MAP()

	void Create( void );

	static void ClearAll( void );

protected:
	LRESULT OnEraseBkgnd( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled );
	LRESULT OnTimer( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled );
	virtual void OnFinalMessage( HWND ) { PostQuitMessage(0); delete this; }

private:
	int m_Time0;
	int m_Duration;
	int m_LastTime;
	HBITMAP m_Bitmap;
	HRGN m_Region;
	RECT m_Rect;

	static std::vector<CMenuFader*> s_Faders;
};
