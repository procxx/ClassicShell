// Classic Shell (c) 2009-2012, Ivo Beltchev
// The sources for Classic Shell are distributed under the MIT open source license

#pragma once

#include "SkinManager.h"
#include "JumpLists.h"
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
	MENU_EMPTY_TOP,
	MENU_RECENT,
	MENU_JUMPITEM,
	MENU_COLUMN_PADDING,
	MENU_COLUMN_BREAK,

	// standard menu items
	MENU_PROGRAMS,
	MENU_APPS,
	MENU_COMPUTER,
	MENU_FAVORITES,
	MENU_DOCUMENTS,
		MENU_USERFILES,
		MENU_USERDOCUMENTS,
		MENU_USERPICTURES,
	MENU_SETTINGS,
		MENU_CONTROLPANEL,
		MENU_CONTROLPANEL_CATEGORIES,
		MENU_NETWORK,
		MENU_SECURITY,
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
	MENU_LOCK,
	MENU_RECENT_ITEMS,
	MENU_SEARCH_BOX,

	MENU_IGNORE=1024, // ignore this item
};

struct StdMenuItem
{
	const wchar_t *command;
	TMenuID id;
	const KNOWNFOLDERID *folder1; // NULL if not used
	const KNOWNFOLDERID *folder2; // NULL if not used

	const wchar_t *label; // localization key
	const wchar_t *tip; // default tooltip
	const wchar_t *iconPath;
	const wchar_t *link;
	unsigned int settings;
	const StdMenuItem *submenu;
	CString labelString, tipString; // additional storage for the strings

	// user settings
	enum
	{
		MENU_OPENUP       = 0x0001, // prefer to open up
		MENU_OPENUP_REC   = 0x0002, // children prefer to open up
		MENU_SORTZA       = 0x0004, // sort backwards
		MENU_SORTZA_REC   = 0x0008, // children sort backwards
		MENU_SORTONCE     = 0x0010, // save the sort order the first time the menu is opened
		MENU_ITEMS_FIRST  = 0x0020, // place the custom items before the folder items
		MENU_TRACK        = 0x0040, // track shortcuts from this menu
		MENU_NOTRACK      = 0x0080, // don't track shortcuts from this menu
		MENU_NOEXPAND     = 0x0100, // don't expand this link item
		MENU_MULTICOLUMN  = 0x0200, // make this item a multi-column item
		MENU_NOEXTENSIONS = 0x0400, // hide extensions
		MENU_INLINE       = 0x0800, // inline sub-items in the parent menu
		MENU_SPLIT_BUTTON = 0x1000, // the item is drawn as a split button

		MENU_NORECENT     = 0x8000  // don't show recent items in the root menu (because a sub-menu uses MENU_RECENT_ITEMS)
	};
};

struct SpecialFolder
{
	const KNOWNFOLDERID *folder;
	unsigned int settings;

	enum
	{
		FOLDER_NOSUBFOLDERS=1, // don't show the subfolders of this folder
		FOLDER_NONEWFOLDER=2, // don't show the "New Folder" command
		FOLDER_NODROP=4, // don't allow reordering, don't show "Sort" and "Auto Arrange" (also implies FOLDER_NONEWFOLDER)
	};
};

extern SpecialFolder g_SpecialFolders[];

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
		MESSAGE_HANDLER( WM_SYSKEYDOWN, OnSysKeyDown )
		MESSAGE_HANDLER( WM_CHAR, OnChar )
		MESSAGE_HANDLER( WM_TIMER, OnTimer )
		MESSAGE_HANDLER( WM_SYSCOMMAND, OnSysCommand )
		MESSAGE_HANDLER( WM_GETOBJECT, OnGetAccObject )
		MESSAGE_HANDLER( WM_CTLCOLOREDIT, OnColorEdit )
		MESSAGE_HANDLER( MCM_REFRESH, OnRefresh )
		MESSAGE_HANDLER( MCM_SETCONTEXTITEM, OnSetContextItem )
		MESSAGE_HANDLER( MCM_REDRAWEDIT, OnRedrawEdit )
		MESSAGE_HANDLER( MCM_REFRESHICONS, OnRefreshIcons )
		MESSAGE_HANDLER( MCM_SETHOTITEM, OnSetHotItem )
		COMMAND_CODE_HANDLER( EN_CHANGE, OnEditChange )
	END_MSG_MAP()

	// options when creating a container
	enum
	{
		CONTAINER_LARGE        = 0x0000001, // use large icons
		CONTAINER_MULTICOLUMN  = 0x0000002, // use multiple columns instead of a single scrolling column
		CONTAINER_MULTICOL_REC = 0x0000004, // the children will be multi-column
		CONTAINER_CONTROLPANEL = 0x0000008, // this is the control panel, don't go into subfolders
		CONTAINER_PROGRAMS     = 0x0000010, // this is a folder from the Start Menu hierarchy (drop operations prefer link over move)
		CONTAINER_DOCUMENTS    = 0x0000020, // sort by time, limit the count (for recent documents)
		CONTAINER_ALLPROGRAMS  = 0x0000040, // this is the main menu of All Programs (combines the Start Menu and Programs folders)
		CONTAINER_RECENT       = 0x0000080, // insert recent programs (sorted by time)
		CONTAINER_LINK         = 0x0000100, // this is an expanded link to a folder (always scrolling)
		CONTAINER_ITEMS_FIRST  = 0x0000200, // put standard items at the top
		CONTAINER_DRAG         = 0x0000400, // allow items to be dragged out
		CONTAINER_DROP         = 0x0000800, // allow dropping of items
		CONTAINER_LEFT         = 0x0001000, // the window is aligned on the left
		CONTAINER_TOP          = 0x0002000, // the window is aligned on the top
		CONTAINER_AUTOSORT     = 0x0004000, // the menu is always in alphabetical order
		CONTAINER_OPENUP_REC   = 0x0008000, // the container's children will prefer to open up instead of down
		CONTAINER_SORTZA       = 0x0010000, // the container will sort backwards by default
		CONTAINER_SORTZA_REC   = 0x0020000, // the container's children will sort backwards by default
		CONTAINER_SORTONCE     = 0x0040000, // the container will save the sort order the first time the menu is opened
		CONTAINER_TRACK        = 0x0080000, // track shortcuts from this menu
		CONTAINER_NOSUBFOLDERS = 0x0100000, // don't go into subfolders
		CONTAINER_NONEWFOLDER  = 0x0200000, // don't show the "New Folder" command
		CONTAINER_SEARCH       = 0x0400000, // this is he search results submenu
		CONTAINER_NOEXTENSIONS = 0x0800000, // hide extensions
		CONTAINER_JUMPLIST     = 0x1000000, // this is a jumplist menu
		CONTAINER_APPS         = 0x2000000, // this is the folder for Metro apps
	};

	CMenuContainer( CMenuContainer *pParent, int index, int options, const StdMenuItem *pStdItem, PIDLIST_ABSOLUTE path1, PIDLIST_ABSOLUTE path2, const CString &regName );
	~CMenuContainer( void );

	void InitItems( void );
	void InitWindow( void );

	static bool CloseStartMenu( void );
	static bool CloseProgramsMenu( void );
	static void HideStartMenu( void );
	static bool IsMenuOpened( void ) { return !s_Menus.empty(); }
	static bool IsMenuWindow( HWND hWnd );
	static bool IgnoreTaskbarTimers( void ) { return !s_Menus.empty() && (s_TaskbarState&ABS_AUTOHIDE); }
	static HWND ToggleStartMenu( HWND startButton, bool bKeyboard, bool bAllPrograms );
	static bool ProcessMouseMessage( HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam );
	static void RefreshIcons( void );

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
	LRESULT OnSysKeyDown( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled );
	LRESULT OnChar( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled );
	LRESULT OnTimer( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled );
	LRESULT OnSysCommand( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled );
	LRESULT OnGetAccObject( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled );
	LRESULT OnSetContextItem( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled );
	LRESULT OnColorEdit( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled );
	LRESULT OnRedrawEdit( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled );
	LRESULT OnRefreshIcons( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled );
	LRESULT OnSetHotItem( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled );
	LRESULT OnEditChange( WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled );
	virtual void OnFinalMessage( HWND ) { Release(); }

private:
	// description of a menu item
	struct MenuItem
	{
		TMenuID id; // if pStdItem!=NULL, this is pStdItem->id. otherwise it can only be MENU_NO, MENU_SEPARATOR, MENU_EMPTY or MENU_EMPTY_TOP
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
		bool bInline:1; // this item is inlined in the parent menu
		bool bInlineFirst:1; // this item is the first from the inlined group
		bool bInlineLast:1; // this item is the last from the inlined group
		bool bSplit:1; // split button item
		bool bJumpList:1; // this item has a jump list
		bool bMetroLink:1; // this is a Windows 8 Metro shortcut
		bool bProtectedLink:1; // this is a protected Metro shortcut (can't be deleted or renamed)
		bool bBlankSeparator:1; // this is a blank separator that is the same size as normal items
		char priority; // used for sorting of the All Programs menu

		// pair of shell items. 2 items are used to combine a user folder with a common folder (I.E. user programs/common programs)
		PIDLIST_ABSOLUTE pItem1;
		PIDLIST_ABSOLUTE pItem2;
		union
		{
			UINT accelerator; // accelerator character, 0 if none
			FILETIME time; // timestamp of the file (for sorting recent documents)
			int jumpIndex; // MAKELONG(group,item)
		};

		bool operator<( const MenuItem &x ) const
		{
			if (priority<x.priority) return true;
			if (priority>x.priority) return false;
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
			return CompareMenuString(name,x.name)<0;
		}

		void SetName( const wchar_t *_name, bool bNoExtensions )
		{
			if (bNoExtensions)
			{
				const wchar_t *end=wcsrchr(_name,'.');
				if (end)
				{
					name=CString(_name,(int)(end-_name));
					return;
				}
			}
			name=_name;
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
			return CompareMenuString(name,x.name)<0;
		}
	};

	// Recent document item (sorts by time, newer first)
	struct Document
	{
		CString name;
		FILETIME time;

		bool operator<( const Document &x ) const { return CompareFileTime(&time,&x.time)>0; }
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
	bool m_bHotArrow;
	CString m_RegName; // name of the registry key to store the item order
	PIDLIST_ABSOLUTE m_Path1a[2];
	PIDLIST_ABSOLUTE m_Path2a[2];
	CComPtr<IShellFolder> m_pDropFoldera[2]; // the primary folder (used only as a drop target)
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

	int m_ClickIndex; // the index of the last clicked item (-2 until the mouse enters the menu for the first time)
	DWORD m_HotPos; // last mouse position over a hot item (used to ignore WM_MOUSEMOVE when the mouse didn't really move)
	int m_HoverItem; // item under the mouse (used for opening a submenu when the mouse hovers over an item)
	int m_ContextItem; // force this to be the hot item while a context menu is up
	HBITMAP m_Bitmap; // the background bitmap
	HBITMAP m_ArrowsBitmap[4]; // normal, selected, normal2, selected2
	HFONT m_Font[4]; // first, second, first separator, second separator
	HRGN m_Region; // the outline region
	int m_MaxWidth;
	bool m_bTwoColumns;
	RECT m_rContent;
	RECT m_rContent2;
	int m_ItemHeight[2];
	int m_MaxItemWidth[2];
	int m_IconTopOffset[2]; // offset from the top of the item to the top of the icon
	int m_TextTopOffset[4]; // offset from the top of the item to the top of the text
	RECT m_rUser1; // the user image (0,0,0,0 if the user image is not shown)
	RECT m_rUser2; // the user name (0,0,0,0 if the user name is not shown)
	RECT m_rPadding; // padding in the menu where right-click is possible

	int m_ScrollCount; // number of items to scroll in the pager
	int m_ScrollHeight; // 0 - don't scroll
	int m_ScrollOffset;
	int m_ScrollButtonSize;
	int m_MouseWheel;
	bool m_bScrollUp, m_bScrollDown;
	bool m_bScrollUpHot, m_bScrollDownHot;
	bool m_bScrollTimer;
	bool m_bNoSearchDraw;
	bool m_bSearchDrawn;
	bool m_bInSearchUpdate;
	bool m_bSearchShowAll;
	int m_SearchIndex;
	CWindow m_SearchBox;
	HBITMAP m_SearchIcons;

	struct SearchItem
	{
		CString name; // all caps (if blank, the item must be ignored)
		PIDLIST_ABSOLUTE pidl;
		int rank;
		mutable int icon;
		bool bMetroLink;

		bool MatchText( const wchar_t *search ) const;
		bool operator<( const SearchItem &item ) const { return rank>item.rank || (rank==item.rank && wcscmp(name,item.name)<0); }
	};

	std::vector<SearchItem> m_SearchItems;
	unsigned int m_ResultsHash;
	void InitItems( const std::vector<SearchItem> &items, const wchar_t *search );

	// additional commands for the context menu
	enum
	{
		CMD_OPEN=1,
		CMD_OPEN_ALL,
		CMD_SORT,
		CMD_AUTOSORT,
		CMD_NEWFOLDER,
		CMD_NEWSHORTCUT,
		CMD_DELETEMRU,
		CMD_DELETEALL,
		CMD_EXPLORE,
		CMD_PIN,

		CMD_LAST,
		CMD_MAX=32767
	};

	// ways to activate a menu item
	enum TActivateType
	{
		ACTIVATE_SELECT, // just selects the item
		ACTIVATE_OPEN, // opens the submenu or selects if not a menu
		ACTIVATE_OPEN_KBD, // same as above, but when done with a keyboard
		ACTIVATE_OPEN_SEARCH, // opens the search results submenu
		ACTIVATE_EXECUTE, // executes the item
		ACTIVATE_MENU, // shows context menu
		ACTIVATE_MENU_BACKGROUND, // shows context menu for the menu itself
		ACTIVATE_RENAME, // renames the item
		ACTIVATE_DELETE, // deletes the item
		ACTIVATE_PROPERTIES, // shows the properties of the item
	};

	// sound events
	enum TMenuSound
	{
		SOUND_MAIN,
		SOUND_POPUP,
		SOUND_COMMAND,
		SOUND_DROP,
	};

	// search state
	enum TSearchState
	{
		SEARCH_NONE, // the search is inactive
		SEARCH_BLANK, // the search box has the focus but is blank
		SEARCH_TEXT, // the search box has the focus and is not blank
		SEARCH_MORE, // there are too many search results
	};

	TSearchState m_SearchState;

	enum
	{
		// timer ID
		TIMER_HOVER=1,
		TIMER_HOVER_SPLIT=2,
		TIMER_SCROLL=3,
		TIMER_TOOLTIP_SHOW=4,
		TIMER_TOOLTIP_HIDE=5,
		TIMER_BALLOON_HIDE=6,

		MCM_REFRESH=WM_USER+10, // posted to force the container to refresh its contents
		MCM_SETCONTEXTITEM=WM_USER+11, // sets the item for the context menu. wParam is the nameHash of the item
		MCM_REDRAWEDIT=WM_USER+12, // redraw the search edit box
		MCM_REFRESHICONS=WM_USER+13, // refreshes the icon list and redraws all menus
		MCM_SETHOTITEM=WM_USER+14, // sets the hot item

		// some constants
		SEPARATOR_HEIGHT=8,
		MIN_SCROLL_HEIGHT=13, // the scroll buttons are at least this tall
		MAX_MENU_ITEMS=500,
		MENU_ANIM_SPEED=200,
		MENU_ANIM_SPEED_SUBMENU=100,
		MENU_FADE_SPEED=400,
		MRU_PROGRAMS_COUNT=20,
		MRU_DEFAULT_COUNT=5,
	};

	void AddFirstFolder( CComPtr<IShellFolder> pFolder, PIDLIST_ABSOLUTE path, std::vector<MenuItem> &items, int options, unsigned int hash0 );
	void AddSecondFolder( CComPtr<IShellFolder> pFolder, PIDLIST_ABSOLUTE path, std::vector<MenuItem> &items, int options, unsigned int hash0 );

	// pPt - optional point in screen space (used only by ACTIVATE_EXECUTE and ACTIVATE_MENU)
	void ActivateItem( int index, TActivateType type, const POINT *pPt, bool bNoModifiers=false );
	void RunUserCommand( bool bPicture );
	void ShowKeyboardCues( void );
	void SetActiveWindow( void );
	void CreateBackground( int width1, int width2, int height1, int height2 ); // width1/2, height1/2 - the first and second content area
	void CreateSubmenuRegion( int width, int height ); // width, height - the content area
	void PostRefreshMessage( void );
	void SaveItemOrder( const std::vector<SortMenuItem> &items );
	void LoadItemOrder( void );
	void AddMRUShortcut( const wchar_t *path );
	void DeleteMRUShortcut( const wchar_t *path );
	void SaveMRUShortcuts( void );
	void LoadMRUShortcuts( void );
	void FadeOutItem( int index );
	bool GetItemRect( int index, RECT &rc );
	int HitTest( const POINT &pt, bool bDrop=false );
	bool DragOut( int index );
	void InvalidateItem( int index );
	void SetHotItem( int index, bool bArrow=false );
	void SetInsertMark( int index, bool bAfter );
	bool GetInsertRect( RECT &rc );
	void DrawBackground( HDC hdc, const RECT &drawRect );
	bool GetDescription( int index, wchar_t *text, int size );
	void UpdateScroll( void );
	void UpdateScroll( const POINT *pt );
	void PlayMenuSound( TMenuSound sound );
	bool CanSelectItem( const MenuItem &item );
	void CollectSearchItems( void );
	void SetSearchState( TSearchState state );
	void UpdateSearchResults( bool bForceShowAll );
	void AddStandardItems( void );
	void UpdateAccelerators( int first, int last );
	void ExecuteCommandElevated( const wchar_t *command );
	void OpenSubMenu( int index, TActivateType type, bool bShift );

	enum
	{
		COLLECT_RECURSIVE =1, // go into subfolders
		COLLECT_PROGRAMS  =2, // only collect programs (.exe, .com, etc)
		COLLECT_FOLDERS   =4, // include folder items
		COLLECT_METRO     =8, // check for metro links (non-recursive)
	};

	void CollectSearchItemsInt( IShellFolder *pFolder, PIDLIST_ABSOLUTE pidl, int flags, int &count );

	static int s_MaxRecentDocuments; // limit for the number of recent documents
	static int s_ScrollMenus; // global scroll menus setting
	static bool s_bRTL; // RTL layout
	static bool s_bKeyboardCues; // show keyboard cues
	static bool s_bOverrideFirstDown; // the first down key from the search box will select the top item
	static bool s_bExpandRight; // prefer expanding submenus to the right
	static bool s_bRecentItems; // show and track recent items
	static bool s_bBehindTaskbar; // the main menu is behind the taskbar (when the taskbar is horizontal)
	static bool s_bShowTopEmpty; // shows the empty item on the top menu so the user can drag items there
	static bool s_bNoDragDrop; // disables drag/drop
	static bool s_bNoContextMenu; // disables the context menu
	static bool s_bExpandLinks; // expand links to folders
	static bool s_bSearchSubWord; // the search will match parts of words
	static bool s_bLogicalSort; // use StrCmpLogical instead of CompareString
	static bool s_bExtensionSort; // sort file names by extension
	static bool s_bAllPrograms; // this is the All Programs menu of the Windows start menu
	static bool s_bNoCommonFolders; // don't show the common folders (start menu and programs)
	static char s_bActiveDirectory; // the Active Directory services are available (-1 - uninitialized)
	static CMenuContainer *s_pDragSource; // the source of the current drag operation
	static bool s_bRightDrag; // dragging with the right mouse button
	static RECT s_MainRect; // area of the main monitor
	static DWORD s_TaskbarState; // the state of the taskbar (ABS_AUTOHIDE and ABS_ALWAYSONTOP)
	static DWORD s_HoverTime;
	static DWORD s_SplitHoverTime;
	static DWORD s_XMouse;
	static DWORD s_SubmenuStyle;
	static CLIPFORMAT s_ShellFormat; // CFSTR_SHELLIDLIST
	static CComPtr<IShellFolder> s_pDesktop; // cached pointer of the desktop object
	static CComPtr<IKnownFolderManager> s_pKnownFolders;
	static HWND s_LastFGWindow; // stores the foreground window to restore later when the menu closes
	static HTHEME s_Theme;
	static HTHEME s_PagerTheme;
	static CWindow s_Tooltip;
	static CWindow s_TooltipBalloon;
	static int s_TipShowTime;
	static int s_TipHideTime;
	static int s_TipShowTimeFolder;
	static int s_TipHideTimeFolder;
	static int s_HotItem;
	static CMenuContainer *s_pHotMenu; // the menu with the hot item
	static int s_TipItem; // the item that needs a tooltip
	static CMenuContainer *s_pTipMenu;

	static std::vector<CMenuContainer*> s_Menus; // all menus, in cascading order
	static volatile HWND s_FirstMenu;
	static std::map<unsigned int,int> s_MenuScrolls; // scroll offset for each sub menu

	static CString s_MRUShortcuts[MRU_PROGRAMS_COUNT];

	struct ItemRank
	{
		unsigned int hash; // hash of the item name in caps
		int rank; // number of times it was used
		bool operator<( const ItemRank &rank ) const { return hash<rank.hash; }
	};

	static std::vector<ItemRank> s_ItemRanks;
	static void SaveItemRanks( void );
	static void LoadItemRanks( void );
	static void AddItemRank( unsigned int hash );

	static wchar_t s_JumpAppId[_MAX_PATH];
	static wchar_t s_JumpAppExe[_MAX_PATH];
	static CJumpList s_JumpList;

	static MenuSkin s_Skin;

	friend class COwnerWindow;
	friend class CMenuAccessible;
	friend LRESULT CALLBACK SubclassTopMenuProc( HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData );

	static int CompareMenuString( const wchar_t *str1, const wchar_t *str2 );
	static void MarginsBlit( HDC hSrc, HDC hDst, const RECT &rSrc, const RECT &rDst, const RECT &rMargins, bool bAlpha );
	static void UpdateUsedIcons( void );
	static LRESULT CALLBACK SubclassSearchBox( HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData );
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

const RECT DEFAULT_SEARCH_PADDING={4,4,4,4};
