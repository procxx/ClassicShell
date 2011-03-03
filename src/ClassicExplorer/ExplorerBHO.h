// Classic Shell (c) 2009-2011, Ivo Beltchev
// The sources for Classic Shell are distributed under the MIT open source license

// ExplorerBHO.h : Declaration of the CExplorerBHO

#pragma once
#include "resource.h"       // main symbols
#include <vector>

#include "ClassicExplorer_i.h"


#if defined(_WIN32_WCE) && !defined(_CE_DCOM) && !defined(_CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA)
#error "Single-threaded COM objects are not properly supported on Windows CE platform, such as the Windows Mobile platforms that do not include full DCOM support. Define _CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA to force ATL to support creating single-thread COM object's and allow use of it's single-threaded COM object implementations. The threading model in your rgs file was set to 'Free' as that is the only threading model supported in non DCOM Windows CE platforms."
#endif



// CExplorerBHO

class ATL_NO_VTABLE CExplorerBHO :
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CExplorerBHO, &CLSID_ExplorerBHO>,
	public IObjectWithSiteImpl<CExplorerBHO>,
	public IDispatchImpl<IExplorerBHO, &IID_IExplorerBHO, &LIBID_ClassicExplorerLib, /*wMajor =*/ 1, /*wMinor =*/ 0>,
	public IDispEventImpl<1,CExplorerBHO,&DIID_DWebBrowserEvents2,&LIBID_SHDocVw,1,1>
{
public:
	CExplorerBHO()
	{
		m_bResetStatus=true;
		m_bForceRefresh=false;
		m_bRemapBands=false;
		m_bNoBreadcrumbs=false;
		m_CurIcon=NULL;
		m_CurPidl=NULL;
		m_NavigatePidl=NULL;
		m_CurPath[0]=0;
		m_Rebar=NULL;
		m_TopWindow=NULL;
		m_Breadcrumbs=NULL;
		m_Hook=m_HookKbd=NULL;
	}

	DECLARE_REGISTRY_RESOURCEID(IDR_EXPLORERBHO)

	BEGIN_SINK_MAP( CExplorerBHO )
		SINK_ENTRY_EX(1, DIID_DWebBrowserEvents2, DISPID_NAVIGATECOMPLETE2, OnNavigateComplete)
		SINK_ENTRY_EX(1, DIID_DWebBrowserEvents2, DISPID_ONQUIT, OnQuit)
	END_SINK_MAP()

	BEGIN_COM_MAP(CExplorerBHO)
		COM_INTERFACE_ENTRY(IExplorerBHO)
		COM_INTERFACE_ENTRY(IObjectWithSite)
		COM_INTERFACE_ENTRY(IDispatch)
	END_COM_MAP()



	DECLARE_PROTECT_FINAL_CONSTRUCT()

	HRESULT FinalConstruct()
	{
		return S_OK;
	}

	void FinalRelease()
	{
	}

	// Navigation pane styles
	enum
	{
		STYLE_CLASSIC=0, // use classic XP style
		STYLE_SIMPLE=1, // use simple XP style
		STYLE_VISTA=2, // no change
	};

	enum
	{
		// from settings
		SPACE_INFOTIP=1, // show the infotip in the status bar if a single item is selected

		// from code
		SPACE_WIN7=2, // running on Win7 (fix the status bar parts and show the disk free space)
	};

public:
	// IObjectWithSite
	STDMETHOD(SetSite)(IUnknown *pUnkSite);

	// DWebBrowserEvents2
	STDMETHOD(OnNavigateComplete)( IDispatch *pDisp, VARIANT *URL );
	STDMETHOD(OnQuit)( void );

private:
	// Super-class the toolbar, so it has a different class name. A program called Folder Menu 3 looks for specific controls in Explorer,
	// and our extra toolbar throws it off
	class CToolbar: public CWindowImpl<CToolbar>
	{
	public:
		DECLARE_WND_SUPERCLASS(L"ClassicShell.UpButton",TOOLBARCLASSNAME);

		BEGIN_MSG_MAP( CToolbar )
		END_MSG_MAP()
	};

	CComPtr<IShellBrowser> m_pBrowser;
	CComPtr<IWebBrowser2> m_pWebBrowser;
	bool m_bResetStatus;
	bool m_bForceRefresh;
	bool m_bNoBreadcrumbs;
	bool m_bRemapBands;
	int m_UpButtonIndex;
	CToolbar m_Toolbar;
	HWND m_TopWindow;
	HWND m_Rebar;
	HICON m_IconNormal, m_IconHot, m_IconPressed, m_IconDisabled;
	HICON m_CurIcon;
	PIDLIST_ABSOLUTE m_CurPidl;
	wchar_t m_CurPath[1024]; // the current path
	CWindow m_ComboBox;
	PIDLIST_ABSOLUTE m_NavigatePidl;
	UINT m_NavigateMsg; // private message that is posted to the progress bar to navigate ti m_NavigatePidl
	HHOOK m_Hook;
	HHOOK m_HookKbd;
	HWND m_Breadcrumbs;
	char m_AltD;

	struct ComboItem
	{
		PIDLIST_ABSOLUTE pidl;
		int indent;
		CString name;
		CString sortName;

		bool operator<( const ComboItem &item ) { return _wcsicmp(sortName,item.sortName)<0; }
	};
	std::vector<ComboItem> m_ComboItems;
	void ClearComboItems( void );

	static int s_AutoNavDelay;

	static LRESULT CALLBACK HookExplorer( int code, WPARAM wParam, LPARAM lParam );
	static LRESULT CALLBACK HookKeyboard( int code, WPARAM wParam, LPARAM lParam );
	static LRESULT CALLBACK SubclassTreeParentProc( HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData );
	static LRESULT CALLBACK SubclassTreeProc( HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData );
	static LRESULT CALLBACK SubclassStatusProc( HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData );
	static LRESULT CALLBACK RebarSubclassProc( HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData );
	static LRESULT CALLBACK BreadcrumbSubclassProc( HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData );
	static LRESULT CALLBACK ProgressSubclassProc( HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData );
};

OBJECT_ENTRY_AUTO(__uuidof(ExplorerBHO), CExplorerBHO)

bool ShowTreeProperties( HWND hwndTree );
