// Classic Shell (c) 2009-2013, Ivo Beltchev
// The sources for Classic Shell are distributed under the MIT open source license

// ClassicCopyExt.h : Declaration of the CClassicCopyExt

#pragma once
#include "resource.h"       // main symbols

#include "ClassicExplorer_i.h"
#include <vector>

#if defined(_WIN32_WCE) && !defined(_CE_DCOM) && !defined(_CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA)
#error "Single-threaded COM objects are not properly supported on Windows CE platform, such as the Windows Mobile platforms that do not include full DCOM support. Define _CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA to force ATL to support creating single-thread COM object's and allow use of it's single-threaded COM object implementations. The threading model in your rgs file was set to 'Free' as that is the only threading model supported in non DCOM Windows CE platforms."
#endif



// CClassicCopyExt

class ATL_NO_VTABLE CClassicCopyExt :
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CClassicCopyExt, &CLSID_ClassicCopyExt>,
	public IShellExtInit,
	public IContextMenu 

{
public:
	CClassicCopyExt()
	{
	}

DECLARE_REGISTRY_RESOURCEID(IDR_CLASSICCOPYEXT)

DECLARE_NOT_AGGREGATABLE(CClassicCopyExt)

BEGIN_COM_MAP(CClassicCopyExt)
	COM_INTERFACE_ENTRY(IShellExtInit)
	COM_INTERFACE_ENTRY(IContextMenu)
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
	// IShellExtInit
	STDMETHODIMP Initialize(PCIDLIST_ABSOLUTE, LPDATAOBJECT, HKEY);

	// IContextMenu
	STDMETHODIMP GetCommandString(UINT_PTR, UINT, UINT*, LPSTR, UINT);
	STDMETHODIMP InvokeCommand(LPCMINVOKECOMMANDINFO);
	STDMETHODIMP QueryContextMenu(HMENU, UINT, UINT, UINT, UINT);
};

OBJECT_ENTRY_AUTO(__uuidof(ClassicCopyExt), CClassicCopyExt)
