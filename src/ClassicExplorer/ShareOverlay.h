// Classic Shell (c) 2009-2010, Ivo Beltchev
// The sources for Classic Shell are distributed under the MIT open source license

// ShareOverlay.h : Declaration of the CShareOverlay

#pragma once
#include "resource.h"       // main symbols

#include "ClassicExplorer_i.h"


#if defined(_WIN32_WCE) && !defined(_CE_DCOM) && !defined(_CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA)
#error "Single-threaded COM objects are not properly supported on Windows CE platform, such as the Windows Mobile platforms that do not include full DCOM support. Define _CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA to force ATL to support creating single-thread COM object's and allow use of it's single-threaded COM object implementations. The threading model in your rgs file was set to 'Free' as that is the only threading model supported in non DCOM Windows CE platforms."
#endif



// CShareOverlay

class ATL_NO_VTABLE CShareOverlay :
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CShareOverlay, &CLSID_ShareOverlay>,
	public IShellIconOverlayIdentifier
{
public:
	CShareOverlay( void );

	DECLARE_REGISTRY_RESOURCEID(IDR_SHAREOVERLAY)

	DECLARE_PROTECT_FINAL_CONSTRUCT()

	HRESULT FinalConstruct( void )
	{
		return S_OK;
	}

	void FinalRelease( void )
	{
	}

public:

	HRESULT _InternalQueryInterface( REFIID iid, void** ppvObject );

	// IShellIconOverlayIdentifier
	STDMETHOD (IsMemberOf)( LPCWSTR pwszPath, DWORD dwAttrib );
	STDMETHOD (GetOverlayInfo)( LPWSTR pwszIconFile, int cchMax, int * pIndex, DWORD * pdwFlags );
	STDMETHOD (GetPriority)( int * pIPriority );

	static void InitOverlay( const wchar_t *icon );

private:
	CComPtr<IShellFolder> m_pDesktop;

	static bool s_bEnabled;
	static int s_Index;
	static wchar_t s_IconPath[_MAX_PATH];
};

OBJECT_ENTRY_AUTO(__uuidof(ShareOverlay), CShareOverlay)
