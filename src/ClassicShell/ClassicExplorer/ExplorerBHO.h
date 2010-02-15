// Classic Shell (c) 2009-2010, Ivo Beltchev
// The sources for Classic Shell are distributed under the MIT open source license

// ExplorerBHO.h : Declaration of the CExplorerBHO

#pragma once
#include "resource.h"       // main symbols

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

	// Options for the folders tree
	enum
	{
		FOLDERS_ALTENTER=1, // enable Alt+Enter support

		FOLDERS_VISTA=0, // no change
		FOLDERS_CLASSIC=2, // use classic XP style
		FOLDERS_SIMPLE=6, // use simple XP style
		FOLDERS_STYLE_MASK=6,

		FOLDERS_NOFADE=8, // don't fade the buttons
		FOLDERS_AUTONAVIGATE=16, // always navigate to selected folder
		FOLDERS_FULLINDENT=32, // use full-size indent

		FOLDERS_DEFAULT=FOLDERS_ALTENTER
	};

	enum
	{
		SPACE_SHOW=1, // show free space and selection size
		SPACE_TOTAL=2, // show total size when nothing is selected
		SPACE_WIN7=4, // running on Win7 (fix the status bar parts and show the disk free space)
	};

public:
	// IObjectWithSite
	STDMETHOD(SetSite)(IUnknown *pUnkSite);

	// DWebBrowserEvents2
	STDMETHOD(OnNavigateComplete)( IDispatch *pDisp, VARIANT *URL );
	STDMETHOD(OnQuit)( void );

private:
	CComPtr<IShellBrowser> m_pBrowser;
	CComPtr<IWebBrowser2> m_pWebBrowser;
	bool m_bResetStatus;
	bool m_bForceRefresh;
	CWindow m_Toolbar;
	HICON m_IconNormal, m_IconHot, m_IconPressed, m_IconDisabled;

	static __declspec(thread) HHOOK s_Hook;

	static LRESULT CALLBACK HookExplorer( int code, WPARAM wParam, LPARAM lParam );
	static LRESULT CALLBACK SubclassStatusProc( HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData );
	static LRESULT CALLBACK RebarSubclassProc( HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData );
};

OBJECT_ENTRY_AUTO(__uuidof(ExplorerBHO), CExplorerBHO)

bool ShowTreeProperties( HWND hwndTree );
