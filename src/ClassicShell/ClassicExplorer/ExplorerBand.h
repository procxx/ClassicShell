// Classic Shell (c) 2009-2010, Ivo Beltchev
// The sources for Classic Shell are distributed under the MIT open source license

// ExplorerBand.h : Declaration of the CExplorerBand

#pragma once
#include "resource.h"       // main symbols
#include "ClassicExplorer_i.h"
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
		ID_RENAME,
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

	BEGIN_MSG_MAP( CBandWindow )
		MESSAGE_HANDLER( WM_CREATE, OnCreate )
		MESSAGE_HANDLER( WM_DESTROY, OnDestroy )
		MESSAGE_HANDLER( WM_CLEAR, OnUpdateUI )
		COMMAND_ID_HANDLER( ID_SETTINGS, OnSettings )
		COMMAND_ID_HANDLER( ID_GOUP, OnNavigate )
		COMMAND_ID_HANDLER( ID_GOBACK, OnNavigate )
		COMMAND_ID_HANDLER( ID_GOFORWARD, OnNavigate )
		COMMAND_ID_HANDLER( ID_EMAIL, OnEmail )
		COMMAND_ID_HANDLER( ID_RENAME, OnRename )
		COMMAND_RANGE_HANDLER( ID_CUT, ID_CUSTOM+100, OnToolbarCommand )
		NOTIFY_CODE_HANDLER( NM_RCLICK, OnRClick )
		NOTIFY_CODE_HANDLER( TBN_GETINFOTIP, OnGetInfoTip )
		NOTIFY_CODE_HANDLER( TBN_DROPDOWN, OnDropDown )
		NOTIFY_CODE_HANDLER( RBN_CHEVRONPUSHED, OnChevron )
	END_MSG_MAP()

	CBandWindow( void ) { m_ImgEnabled=m_ImgDisabled=NULL; }

	HWND GetToolbar( void ) { return m_Toolbar.m_hWnd; }
	void SetBrowser( IShellBrowser *pBrowser ) { m_pBrowser=pBrowser; }
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
	LRESULT OnNavigate( WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled );
	LRESULT OnToolbarCommand( WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled );
	LRESULT OnEmail( WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled );
	LRESULT OnRename( WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled );
	LRESULT OnSettings( WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled );
	LRESULT OnRClick( int idCtrl, LPNMHDR pnmh, BOOL& bHandled );
	LRESULT OnGetInfoTip( int idCtrl, LPNMHDR pnmh, BOOL& bHandled );
	LRESULT OnDropDown( int idCtrl, LPNMHDR pnmh, BOOL& bHandled );
	LRESULT OnChevron( int idCtrl, LPNMHDR pnmh, BOOL& bHandled );

private:
	CWindow m_Toolbar;
	CComPtr<IShellBrowser> m_pBrowser;
	HIMAGELIST m_ImgEnabled;
	HIMAGELIST m_ImgDisabled;
	int m_MenuIconSize;

	struct StdToolbarItem
	{
		int id;
		const char *tipKey; // localization key for the tooltip
		const wchar_t *tip; // default tooltip
		int icon; // index in shell32.dll

		const wchar_t *name; // default name
		const wchar_t *command;
		const wchar_t *link;
		const wchar_t *iconPath;
		const wchar_t *iconPathD;
		CString regName; // name of the registry value to check for enabled/checked state

		const StdToolbarItem *submenu;
		mutable HBITMAP menuIcon;
		mutable HBITMAP menuIconD;
		mutable bool bIconLoaded;

		bool bDisabled;
		bool bChecked;
	};

	static const StdToolbarItem s_StdItems[];

	std::vector<StdToolbarItem> m_Items;
	void ParseToolbar( DWORD stdEnabled );
	void SendShellTabCommand( int command );
	HMENU CreateDropMenu( const StdToolbarItem *pItem );
	HMENU CreateDropMenuRec( const StdToolbarItem *pItem, std::vector<HMODULE> &modules, HMODULE hShell32 );

	static void ParseToolbarItem( const wchar_t *name, StdToolbarItem &item );
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

	static LRESULT CALLBACK RebarSubclassProc( HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData );
	static LRESULT CALLBACK ParentSubclassProc( HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData );
};

OBJECT_ENTRY_AUTO(__uuidof(ExplorerBand), CExplorerBand)
