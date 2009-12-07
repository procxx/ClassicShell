// Classic Shell (c) 2009, Ivo Beltchev
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
	public IDispatchImpl<IExplorerBHO, &IID_IExplorerBHO, &LIBID_ClassicExplorerLib, /*wMajor =*/ 1, /*wMinor =*/ 0>
{
public:
	CExplorerBHO()
	{
		s_hwndTree=0;
	}

DECLARE_REGISTRY_RESOURCEID(IDR_EXPLORERBHO)


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

public:
	// IObjectWithSite
	STDMETHOD(SetSite)(IUnknown *pUnkSite);

private:
	static __declspec(thread) HHOOK s_Hook;
	static __declspec(thread) HWND s_hwndTree;

	static LRESULT CALLBACK HookExplorer( int code, WPARAM wParam, LPARAM lParam );
};

OBJECT_ENTRY_AUTO(__uuidof(ExplorerBHO), CExplorerBHO)
