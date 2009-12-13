// Classic Shell (c) 2009, Ivo Beltchev
// The sources for Classic Shell are distributed under the MIT open source license

// ExplorerBand.h : Declaration of the CExplorerBand

#pragma once
#include "resource.h"       // main symbols
#include "ClassicExplorer_i.h"


class CBandWindow: public CWindowImpl<CBandWindow>
{
public:

	enum
	{
		// toolbar command IDs
		ID_SETTINGS=1,
		ID_GOUP,
		ID_CUT,
		ID_COPY,
		ID_PASTE,
		ID_DELETE,
		ID_PROPERTIES,

		ID_LAST
	};

	DECLARE_WND_CLASS(L"ClassicExplorer.CBandWindow")

	BEGIN_MSG_MAP( CBandWindow )
		MESSAGE_HANDLER( WM_CREATE, OnCreate )
		MESSAGE_HANDLER( WM_DESTROY, OnDestroy )
		COMMAND_ID_HANDLER( ID_GOUP, OnGoUp )
		COMMAND_ID_HANDLER( ID_CUT, OnFileOperation )
		COMMAND_ID_HANDLER( ID_COPY, OnFileOperation )
		COMMAND_ID_HANDLER( ID_PASTE, OnFileOperation )
		COMMAND_ID_HANDLER( ID_DELETE, OnFileOperation )
		COMMAND_ID_HANDLER( ID_PROPERTIES, OnFileOperation )
		COMMAND_ID_HANDLER( ID_SETTINGS, OnSettings )
	END_MSG_MAP()

	CBandWindow( void ) { m_Enabled=NULL; }

	HWND GetToolbar( void ) { return m_Toolbar.m_hWnd; }
	void SetBrowser( IShellBrowser *pBrowser ) { m_pBrowser=pBrowser; }
	void UpdateToolbar( void );

protected:
	// Handler prototypes:
	//  LRESULT MessageHandler(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
	//  LRESULT CommandHandler(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled);
	//  LRESULT NotifyHandler(int idCtrl, LPNMHDR pnmh, BOOL& bHandled);
	LRESULT OnCreate( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled );
	LRESULT OnDestroy( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled );
	LRESULT OnGoUp( WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled );
	LRESULT OnFileOperation( WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled );
	LRESULT OnSettings( WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled );

private:
	CWindow m_Toolbar;
	CComPtr<IShellBrowser> m_pBrowser;
	HIMAGELIST m_Enabled;
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
		SINK_ENTRY_EX(1, DIID_DWebBrowserEvents2, DISPID_DOWNLOADCOMPLETE, OnDownloadComplete)
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
	STDMETHOD(OnDownloadComplete)( void );
	STDMETHOD(OnQuit)( void );

protected:
	bool m_bSubclassRebar; // the rebar needs subclassing
	bool m_bSubclassedRebar; // the rebar is subclassed
	bool m_bBandNewLine; // our band is on a new line (has RBBS_BREAK style)
	CBandWindow m_BandWindow;
	CComPtr<IWebBrowser2> m_pWebBrowser;

	static LRESULT CALLBACK RebarSubclassProc( HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData );
};

OBJECT_ENTRY_AUTO(__uuidof(ExplorerBand), CExplorerBand)
