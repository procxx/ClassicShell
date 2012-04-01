// Classic Shell (c) 2009-2012, Ivo Beltchev
// The sources for Classic Shell are distributed under the MIT open source license

#pragma once

#if defined(_WIN32_WCE) && !defined(_CE_DCOM) && !defined(_CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA)
#error "Single-threaded COM objects are not properly supported on Windows CE platform, such as the Windows Mobile platforms that do not include full DCOM support. Define _CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA to force ATL to support creating single-thread COM object's and allow use of it's single-threaded COM object implementations. The threading model in your rgs file was set to 'Free' as that is the only threading model supported in non DCOM Windows CE platforms."
#endif

#include <exdispid.h>
#include <shobjidl.h>
#include <map>

// CClassicIE9BHO

class ATL_NO_VTABLE CClassicIE9BHO :
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CClassicIE9BHO, &CLSID_ClassicIE9BHO>,
	public IObjectWithSiteImpl<CClassicIE9BHO>,
	public IDispEventImpl<1,CClassicIE9BHO,&DIID_DWebBrowserEvents2,&LIBID_SHDocVw,1,1>
{
public:
	CClassicIE9BHO()
	{
		m_StatusBar=NULL;
		m_Tooltip=NULL;
		m_ProgressBar=NULL;
		m_Progress=-1;
	}

	static HRESULT WINAPI UpdateRegistry( BOOL bRegister );

	BEGIN_SINK_MAP( CClassicIE9BHO )
		SINK_ENTRY_EX(1, DIID_DWebBrowserEvents2, DISPID_NAVIGATECOMPLETE2, OnNavigateComplete)
		SINK_ENTRY_EX(1, DIID_DWebBrowserEvents2, DISPID_PROGRESSCHANGE, OnProgressChange)
		SINK_ENTRY_EX(1, DIID_DWebBrowserEvents2, DISPID_ONQUIT, OnQuit)
	END_SINK_MAP()

	BEGIN_COM_MAP(CClassicIE9BHO)
		COM_INTERFACE_ENTRY(IObjectWithSite)
	END_COM_MAP()



	DECLARE_PROTECT_FINAL_CONSTRUCT()

	HRESULT FinalConstruct()
	{
		return S_OK;
	}

	// IObjectWithSite
	STDMETHOD(SetSite)(IUnknown *pUnkSite);

	// DWebBrowserEvents2
	STDMETHOD(OnNavigateComplete)( IDispatch *pDisp, VARIANT *URL );
	STDMETHOD(OnProgressChange)( long progress, long progressMax );
	STDMETHOD(OnQuit)( void );

private:
	enum
	{
		PART_TEXT,
		PART_PROGRESS,
		PART_ZONE,
		PART_ZOOM,

		PART_COUNT,
		PART_OFFSET=100,

		PROGRESS_WIDTH=110,
		MIN_TEXT_WIDTH=100,
	};

	CComPtr<IShellBrowser>m_pBrowser;
	CComPtr<IWebBrowser2> m_pWebBrowser;
	CComPtr<IInternetZoneManager> m_pZoneManager;
	CComPtr<IInternetSecurityManager> m_pSecurityManager;

	HWND m_StatusBar;
	HWND m_Tooltip;
	HWND m_ProgressBar;
	CString m_ProtectedMode;
	int m_TextWidth;
	int m_Progress;
	std::map<unsigned int,HICON> m_IconCache;

	static LRESULT CALLBACK SubclassStatusProc( HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData );

	void ResetParts( void );
};

OBJECT_ENTRY_AUTO(__uuidof(ClassicIE9BHO), CClassicIE9BHO)
