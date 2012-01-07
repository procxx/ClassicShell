// Classic Shell (c) 2009-2011, Ivo Beltchev
// The sources for Classic Shell are distributed under the MIT open source license

// ExplorerBand.h : Declaration of the CExplorerBand

#pragma once
#include "resource.h"
#include "ClassicExplorer_i.h"
#include "SettingsParser.h"
#include <vector>

class CBandWindow: public CWindowImpl<CBandWindow>
{
public:

	enum
	{
		ID_LAST=-1,
		ID_SEPARATOR=0,

		// standard toolbar commands
		ID_SETTINGS=1,
		ID_GOUP,
		ID_CUT,
		ID_COPY,
		ID_PASTE,
		ID_PASTE_SHORTCUT,
		ID_DELETE,
		ID_PROPERTIES,
		ID_EMAIL,

		ID_LAST_STD, // last standard command

		// additional supported commands
		ID_MOVETO,
		ID_COPYTO,
		ID_UNDO,
		ID_REDO,
		ID_SELECTALL,
		ID_DESELECT,
		ID_INVERT,
		ID_GOBACK,
		ID_GOFORWARD,
		ID_REFRESH,
		ID_STOP,
		ID_RENAME,
		ID_NEWFOLDER,
		ID_VIEW_TILES,
		ID_VIEW_DETAILS,
		ID_VIEW_LIST,
		ID_VIEW_CONTENT,
		ID_VIEW_ICONS1,
		ID_VIEW_ICONS2,
		ID_VIEW_ICONS3,
		ID_VIEW_ICONS4,

		ID_CUSTOM=100,
	};

	DECLARE_WND_CLASS(L"ClassicShell.CBandWindow")

	enum { BWM_UPDATEBUTTONS=WM_USER };

	BEGIN_MSG_MAP( CBandWindow )
		MESSAGE_HANDLER( WM_CREATE, OnCreate )
		MESSAGE_HANDLER( WM_DESTROY, OnDestroy )
		MESSAGE_HANDLER( WM_CLEAR, OnUpdateUI )
		MESSAGE_HANDLER( WM_COMMAND, OnCommand )
		MESSAGE_HANDLER( BWM_UPDATEBUTTONS, OnUpdateButtons )
		NOTIFY_CODE_HANDLER( NM_RCLICK, OnRClick )
		NOTIFY_CODE_HANDLER( TBN_GETINFOTIP, OnGetInfoTip )
		NOTIFY_CODE_HANDLER( TBN_DROPDOWN, OnDropDown )
		NOTIFY_CODE_HANDLER( RBN_CHEVRONPUSHED, OnChevron )
	END_MSG_MAP()

	CBandWindow( void ) { m_ImgEnabled=m_ImgDisabled=NULL; }

	HWND GetToolbar( void ) { return m_Toolbar.m_hWnd; }
	void SetBrowsers( IShellBrowser *pBrowser, IWebBrowser2 *pWebBrowser ) { m_pBrowser=pBrowser; m_pWebBrowser=pWebBrowser; }
	void UpdateToolbar( void );
	void EnableButton( int cmd, bool bEnable );

protected:
	// Handler prototypes:
	//  LRESULT MessageHandler(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
	//  LRESULT CommandHandler(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled);
	//  LRESULT NotifyHandler(int idCtrl, LPNMHDR pnmh, BOOL& bHandled);
	LRESULT OnCreate( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled );
	LRESULT OnDestroy( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled );
	LRESULT OnUpdateUI( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled );
	LRESULT OnCommand( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled );
	LRESULT OnUpdateButtons( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled );
	LRESULT OnRClick( int idCtrl, LPNMHDR pnmh, BOOL& bHandled );
	LRESULT OnGetInfoTip( int idCtrl, LPNMHDR pnmh, BOOL& bHandled );
	LRESULT OnDropDown( int idCtrl, LPNMHDR pnmh, BOOL& bHandled );
	LRESULT OnChevron( int idCtrl, LPNMHDR pnmh, BOOL& bHandled );

private:
	CWindow m_Toolbar;
	CComPtr<IShellBrowser> m_pBrowser;
	CComPtr<IWebBrowser2> m_pWebBrowser;
	HIMAGELIST m_ImgEnabled;
	HIMAGELIST m_ImgDisabled;
	int m_MenuIconSize;

	struct StdToolbarItem
	{
		int id;
		const wchar_t *command;
		const wchar_t *link;
		const wchar_t *label; // text on the button
		const wchar_t *tip; // default tooltip
		const wchar_t *iconPath;
		const wchar_t *iconPathD;
		CString regName; // name of the registry value to check for enabled/checked state
		CString labelString, tipString; // additional storage for the strings
		const StdToolbarItem *submenu;
		mutable HBITMAP menuIcon;
		mutable HBITMAP menuIconD;
		mutable CString menuText;
		mutable bool bIconLoaded; // the menu icon is loaded
		bool bDisabled;
		bool bChecked;
	};

	std::vector<StdToolbarItem> m_Items;
	std::vector<TBBUTTON> m_Buttons;
	CSettingsParser m_Parser;

	void ParseToolbar( void );
	void ParseToolbarItem( const wchar_t *name, StdToolbarItem &item );
	void SendShellTabCommand( int command );
	HMENU CreateDropMenu( const StdToolbarItem *pItem );
	HMENU CreateDropMenuRec( const StdToolbarItem *pItem, std::vector<HMODULE> &modules, HMODULE hShell32 );
	void SendEmail( void );
	void NewFolder( void );
	void ExecuteCommandFile( const wchar_t *pText );
	void ExecuteCustomCommand( const wchar_t *pCommand );
	void ViewByProperty( IFolderView2 *pView, const wchar_t *pProperty, bool bGroup );

	static LRESULT CALLBACK ToolbarSubclassProc( HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData );
};


// CExplorerBand

class ATL_NO_VTABLE CExplorerBand :
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CExplorerBand,&CLSID_ExplorerBand>,
	public IObjectWithSiteImpl<CExplorerBand>,
	public IDeskBand,
	public IDispEventImpl<1,CExplorerBand,&DIID_DWebBrowserEvents2,&LIBID_SHDocVw,1,1>
{
public:
	CExplorerBand( void );

	DECLARE_REGISTRY_RESOURCEID(IDR_EXPLORERBAND)

	BEGIN_SINK_MAP( CExplorerBand )
		SINK_ENTRY_EX(1, DIID_DWebBrowserEvents2, DISPID_NAVIGATECOMPLETE2, OnNavigateComplete)
		SINK_ENTRY_EX(1, DIID_DWebBrowserEvents2, DISPID_COMMANDSTATECHANGE, OnCommandStateChange)
		SINK_ENTRY_EX(1, DIID_DWebBrowserEvents2, DISPID_ONQUIT, OnQuit)
	END_SINK_MAP()

	BEGIN_COM_MAP(CExplorerBand)
		COM_INTERFACE_ENTRY( IOleWindow )
		COM_INTERFACE_ENTRY( IObjectWithSite )
		COM_INTERFACE_ENTRY_IID( IID_IDockingWindow, IDockingWindow )
		COM_INTERFACE_ENTRY_IID( IID_IDeskBand, IDeskBand )
	END_COM_MAP()



	DECLARE_PROTECT_FINAL_CONSTRUCT()

	HRESULT FinalConstruct()
	{
		return S_OK;
	}

	void FinalRelease()
	{
	}

public:

	// IDeskBand
	STDMETHOD(GetBandInfo)( DWORD dwBandID, DWORD dwViewMode, DESKBANDINFO* pdbi );

	// IObjectWithSite
	STDMETHOD(SetSite)( IUnknown* pUnkSite );

	// IOleWindow
	STDMETHOD(GetWindow)( HWND* phwnd );
	STDMETHOD(ContextSensitiveHelp)( BOOL fEnterMode );

	// IDockingWindow
	STDMETHOD(CloseDW)( unsigned long dwReserved );
	STDMETHOD(ResizeBorderDW)( const RECT* prcBorder, IUnknown* punkToolbarSite, BOOL fReserved );
	STDMETHOD(ShowDW)( BOOL fShow );

	// DWebBrowserEvents2
	STDMETHOD(OnNavigateComplete)( IDispatch *pDisp, VARIANT *URL );
	STDMETHOD(OnCommandStateChange)( long Command, VARIANT_BOOL Enable );
	STDMETHOD(OnQuit)( void );

protected:
	bool m_bSubclassRebar; // the rebar needs subclassing
	bool m_bSubclassedRebar; // the rebar is subclassed
	bool m_bBandNewLine; // our band is on a new line (has RBBS_BREAK style)
	bool m_bHandleSetInfo; // mess with the RB_SETBANDINFO message
	CBandWindow m_BandWindow;
	CComPtr<IWebBrowser2> m_pWebBrowser;
	HWND m_TopWindow;

	static LRESULT CALLBACK RebarSubclassProc( HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData );
	static LRESULT CALLBACK ParentSubclassProc( HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData );
};

OBJECT_ENTRY_AUTO(__uuidof(ExplorerBand), CExplorerBand)
