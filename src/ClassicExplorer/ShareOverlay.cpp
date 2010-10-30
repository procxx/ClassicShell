// Classic Shell (c) 2009-2010, Ivo Beltchev
// The sources for Classic Shell are distributed under the MIT open source license

// ShareOverlay.cpp : Implementation of CShareOverlay

#include "stdafx.h"
#include "ShareOverlay.h"

// CShareOverlay - adds an overlay icon to the shared folders

bool CShareOverlay::s_bEnabled=false;
int CShareOverlay::s_Index;
wchar_t CShareOverlay::s_IconPath[_MAX_PATH];

CShareOverlay::CShareOverlay( void )
{
	SHGetDesktopFolder(&m_pDesktop);
}

void CShareOverlay::InitOverlay( const wchar_t *icon )
{
	s_bEnabled=true;
	if (icon)
	{
		Strcpy(s_IconPath,_countof(s_IconPath),icon);
		wchar_t *c=wcsrchr(s_IconPath,',');
		if (c)
		{
			*c=0;
			s_Index=-_wtol(c+1);
		}
		else
			s_Index=0;
	}
	else
	{
		Strcpy(s_IconPath,_countof(s_IconPath),L"%windir%\\system32\\imageres.dll");
		s_Index=-164;
	}
	DoEnvironmentSubst(s_IconPath,_countof(s_IconPath));
}

HRESULT CShareOverlay::_InternalQueryInterface( REFIID iid, void** ppvObject )
{
	if (iid==IID_IUnknown)
	{
		AddRef();
		*ppvObject=static_cast<IUnknown*>(this);
		return S_OK;
	}
	if (iid==IID_IShellIconOverlayIdentifier && s_bEnabled)
	{
		// only support IShellIconOverlayIdentifier if s_bEnabled is true
		AddRef();
		*ppvObject=static_cast<IShellIconOverlayIdentifier*>(this);
		return S_OK;
	}
	*ppvObject=NULL;
	return E_NOINTERFACE;
}

STDMETHODIMP CShareOverlay::IsMemberOf( LPCWSTR pwszPath, DWORD dwAttrib )
{
	// must use IShellFolder::GetAttributesOf to get the correct attributes instead of SHGetFileInfo or IShellFolder::ParseDisplayName
	// SHGetFileInfo gives the wrong result for some system folders like %userprofile%\Desktop (on Windows7 and Vista)
	// IShellFolder::ParseDisplayName returns the wrong attributes for the contents of the Recycle Bin (on Windows7 only)
	PIDLIST_RELATIVE pidl=NULL;
	HRESULT res=S_FALSE;
	if (m_pDesktop)
	{
		if (SUCCEEDED(m_pDesktop->ParseDisplayName(NULL,NULL,(LPWSTR)pwszPath,NULL,&pidl,NULL)))
		{
			CComPtr<IShellFolder> pFolder;
			PCUITEMID_CHILD child;
			if (SUCCEEDED(SHBindToParent(pidl,IID_IShellFolder,(void**)&pFolder,&child)))
			{
				ULONG attrib=SFGAO_SHARE;
				if (SUCCEEDED(pFolder->GetAttributesOf(1,&child,&attrib)) && (attrib&SFGAO_SHARE))
					res=S_OK;
			}
			ILFree(pidl);
		}
	}
	return res;
}

STDMETHODIMP CShareOverlay::GetOverlayInfo( LPWSTR pwszIconFile, int cchMax, int * pIndex, DWORD * pdwFlags )
{
	Strcpy(pwszIconFile,cchMax,s_IconPath);
	*pIndex=s_Index;
	*pdwFlags=ISIOI_ICONFILE;
	if (s_Index)
		*pdwFlags|=ISIOI_ICONINDEX;
	return S_OK;
}

STDMETHODIMP CShareOverlay::GetPriority( int * pIPriority )
{
	*pIPriority=0;
	return S_OK;
}
