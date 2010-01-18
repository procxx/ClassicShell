// Classic Shell (c) 2009-2010, Ivo Beltchev
// The sources for Classic Shell are distributed under the MIT open source license

#include "stdafx.h"
#include "Accessibility.h"
#include "MenuContainer.h"
#include "TranslationSettings.h"

CMenuAccessible::CMenuAccessible( CMenuContainer *pOwner )
{
	m_RefCount=1;
	m_pOwner=pOwner;
	CreateStdAccessibleObject(pOwner->m_hWnd,OBJID_CLIENT,IID_IAccessible,(void**)&m_pStdAccessible);
}

CMenuAccessible::~CMenuAccessible( void )
{
}

void CMenuAccessible::Reset( void )
{
	m_pOwner=NULL;
	m_pStdAccessible=NULL;
}

HRESULT STDMETHODCALLTYPE CMenuAccessible::get_accParent( IDispatch **ppdispParent )
{
	if (!m_pOwner) return RPC_E_DISCONNECTED;
	if (m_pStdAccessible)
		return m_pStdAccessible->get_accParent(ppdispParent);
	*ppdispParent=NULL;
	return S_FALSE;
}

HRESULT STDMETHODCALLTYPE CMenuAccessible::get_accChildCount( long *pcountChildren )
{
	if (!m_pOwner) return RPC_E_DISCONNECTED;
	*pcountChildren=(long)m_pOwner->m_Items.size();
	return S_OK;
}

HRESULT STDMETHODCALLTYPE CMenuAccessible::get_accChild( VARIANT varChild, IDispatch **ppdispChild )
{
	if (!m_pOwner) return RPC_E_DISCONNECTED;
	*ppdispChild=NULL; // no child IAccessibles
	if (varChild.vt!=VT_I4) return E_INVALIDARG;
	return S_FALSE;
}

HRESULT STDMETHODCALLTYPE CMenuAccessible::get_accName( VARIANT varChild, BSTR *pszName )
{
	if (!m_pOwner) return RPC_E_DISCONNECTED;
	*pszName=NULL;
	if (varChild.vt!=VT_I4) return S_FALSE;
	if (varChild.lVal==CHILDID_SELF) return S_FALSE;
	int index=varChild.lVal-1;
	if (index<0 || index>=(int)m_pOwner->m_Items.size()) return S_FALSE;
	if (m_pOwner->m_Items[index].id==MENU_SEPARATOR) return S_FALSE;
	wchar_t text[256];
	wcscpy_s(text,m_pOwner->m_Items[index].name);
	for (wchar_t *c1=text,*c2=text;;c1++)
	{
		if (*c1!='&')
			*c2++=*c1;
		if (*c1==0) break;
	}
	*pszName=SysAllocString(text);
	return S_OK;
}

HRESULT STDMETHODCALLTYPE CMenuAccessible::get_accDescription( VARIANT varChild, BSTR *pszDescription )
{
	return get_accName(varChild,pszDescription);
}

HRESULT STDMETHODCALLTYPE CMenuAccessible::get_accRole( VARIANT varChild, VARIANT *pvarRole )
{
	if (!m_pOwner) return RPC_E_DISCONNECTED;
	pvarRole->vt=VT_EMPTY;
	if (varChild.vt!=VT_I4) return E_INVALIDARG;
	if (varChild.lVal==CHILDID_SELF)
	{
		pvarRole->vt=VT_I4;
		pvarRole->lVal=ROLE_SYSTEM_MENUPOPUP;
		return S_OK;
	}
	int index=varChild.lVal-1;
	if (index<0 || index>=(int)m_pOwner->m_Items.size()) return E_INVALIDARG;
	pvarRole->vt=VT_I4;
	pvarRole->lVal=m_pOwner->m_Items[index].id==MENU_SEPARATOR?ROLE_SYSTEM_SEPARATOR:ROLE_SYSTEM_MENUITEM;
	return S_OK;
}

HRESULT STDMETHODCALLTYPE CMenuAccessible::get_accState( VARIANT varChild, VARIANT *pvarState )
{
	if (!m_pOwner) return RPC_E_DISCONNECTED;
	pvarState->vt=VT_EMPTY;
	if (varChild.vt!=VT_I4) return E_INVALIDARG;
	int flags=STATE_SYSTEM_FOCUSABLE;
	int index=varChild.lVal-1;
	if (index>=0 && index<(int)m_pOwner->m_Items.size())
	{
		const CMenuContainer::MenuItem &item=m_pOwner->m_Items[index];
		CWindow toolbar=m_pOwner->m_Toolbars[item.column];
		if (toolbar.SendMessage(TB_GETHOTITEM)==item.btnIndex)
			flags|=STATE_SYSTEM_FOCUSED;
		if (item.bFolder)
			flags|=STATE_SYSTEM_HASPOPUP;
		if (item.id==MENU_SEPARATOR)
			flags=0;
	}
	pvarState->vt=VT_I4;
	pvarState->lVal=flags;
	return S_OK;
}

HRESULT STDMETHODCALLTYPE CMenuAccessible::get_accKeyboardShortcut( VARIANT varChild, BSTR *pszKeyboardShortcut )
{
	if (!m_pOwner) return RPC_E_DISCONNECTED;
	*pszKeyboardShortcut=NULL;
	if (varChild.vt!=VT_I4) return E_INVALIDARG;
	int flags=0;
	int index=varChild.lVal-1;
	if (index<0 || index>=(int)m_pOwner->m_Items.size())
		return S_FALSE;
	const CMenuContainer::MenuItem &item=m_pOwner->m_Items[index];
	wchar_t str[2]={item.accelerator,0};
	*pszKeyboardShortcut=SysAllocString(str);
	return S_OK;
}

HRESULT STDMETHODCALLTYPE CMenuAccessible::get_accFocus( VARIANT *pvarChild )
{
	if (!m_pOwner) return RPC_E_DISCONNECTED;
	HWND focus=GetFocus();
	pvarChild->vt=VT_EMPTY;
	for (std::vector<CWindow>::iterator it=m_pOwner->m_Toolbars.begin();it!=m_pOwner->m_Toolbars.end();++it)
	{
		if (it->m_hWnd==focus)
		{
			int index=(int)it->SendMessage(TB_GETHOTITEM);
			if (index>=0)
			{
				TBBUTTON button;
				it->SendMessage(TB_GETBUTTON,index,(LPARAM)&button);
				pvarChild->vt=VT_I4;
				pvarChild->lVal=(int)button.dwData-CMenuContainer::ID_OFFSET+1;
			}
			break;
		}
	}
	return S_OK;
}

HRESULT STDMETHODCALLTYPE CMenuAccessible::get_accSelection( VARIANT *pvarChildren )
{
	if (!m_pOwner) return RPC_E_DISCONNECTED;
	pvarChildren->vt=VT_EMPTY;
	return S_OK;
}

HRESULT STDMETHODCALLTYPE CMenuAccessible::get_accDefaultAction( VARIANT varChild, BSTR *pszDefaultAction )
{
	if (!m_pOwner) return RPC_E_DISCONNECTED;
	*pszDefaultAction=NULL;
	if (varChild.vt!=VT_I4) return E_INVALIDARG;
	if (varChild.lVal==CHILDID_SELF)
	{
		*pszDefaultAction=SysAllocString(FindTranslation("Menu.ActionClose",L"Close"));
		return S_OK;
	}
	int index=varChild.lVal-1;
	if (index<0 || index>=(int)m_pOwner->m_Items.size())
		return S_FALSE;
	const CMenuContainer::MenuItem &item=m_pOwner->m_Items[index];
	if (item.id!=MENU_SEPARATOR && item.id!=MENU_EMPTY)
		*pszDefaultAction=SysAllocString(item.bFolder?FindTranslation("Menu.ActionOpen",L"Open"):FindTranslation("Menu.ActionExecute",L"Execute"));
	return S_OK;
}

HRESULT STDMETHODCALLTYPE CMenuAccessible::accSelect( long flagsSelect, VARIANT varChild )
{
	if (!m_pOwner) return RPC_E_DISCONNECTED;
	if (varChild.vt!=VT_I4) return E_INVALIDARG;
	if (flagsSelect&SELFLAG_TAKEFOCUS)
	{
		int index=varChild.lVal-1;
		if (index<0 || index>=(int)m_pOwner->m_Items.size())
			return S_FALSE;
		m_pOwner->ActivateItem(index,CMenuContainer::ACTIVATE_SELECT,NULL);
	}
	return S_OK;
}

HRESULT STDMETHODCALLTYPE CMenuAccessible::accLocation( long *pxLeft, long *pyTop, long *pcxWidth, long *pcyHeight, VARIANT varChild )
{
	if (!m_pOwner) return RPC_E_DISCONNECTED;
	if (varChild.vt!=VT_I4) return E_INVALIDARG;
	RECT rc;
	if (varChild.lVal==CHILDID_SELF)
	{
		m_pOwner->GetWindowRect(&rc);
	}
	else
	{
		int index=varChild.lVal-1;
		if (index<0 || index>=(int)m_pOwner->m_Items.size())
			return S_FALSE;
		const CMenuContainer::MenuItem &item=m_pOwner->m_Items[index];
		CWindow toolbar=m_pOwner->m_Toolbars[item.column];
		toolbar.SendMessage(TB_GETITEMRECT,item.btnIndex,(LPARAM)&rc);
		toolbar.ClientToScreen(&rc);
	}
	*pxLeft=rc.left;
	*pyTop=rc.top;
	*pcxWidth=rc.right-rc.left;
	*pcyHeight=rc.bottom-rc.top;
	return S_OK;
}

HRESULT STDMETHODCALLTYPE CMenuAccessible::accNavigate( long navDir, VARIANT varStart, VARIANT *pvarEndUpAt )
{
	if (!m_pOwner) return RPC_E_DISCONNECTED;
	pvarEndUpAt->vt=VT_EMPTY;
	if (varStart.vt!=VT_I4) return E_INVALIDARG;

	switch (navDir)
	{
	case NAVDIR_FIRSTCHILD:
		if (varStart.lVal!=CHILDID_SELF) return S_FALSE;
		pvarEndUpAt->vt=VT_I4;
		pvarEndUpAt->lVal=1;
		break;

	case NAVDIR_LASTCHILD:
		if (varStart.lVal!=CHILDID_SELF) return S_FALSE;
		pvarEndUpAt->vt=VT_I4;
		pvarEndUpAt->lVal=(int)m_pOwner->m_Items.size();
		break;

	case NAVDIR_NEXT:   
	case NAVDIR_DOWN:
		if (varStart.lVal==CHILDID_SELF)
		{
			if (m_pStdAccessible)
				return m_pStdAccessible->accNavigate(navDir,varStart,pvarEndUpAt);
			return S_FALSE;
		}
		if (varStart.lVal>=(int)m_pOwner->m_Items.size())
			pvarEndUpAt->vt=VT_EMPTY;
		else
		{
			pvarEndUpAt->vt=VT_I4;
			pvarEndUpAt->lVal=varStart.lVal+1;
		}
		break;

	case NAVDIR_PREVIOUS:
	case NAVDIR_UP:
		if (varStart.lVal==CHILDID_SELF)
		{
			if (m_pStdAccessible)
				return m_pStdAccessible->accNavigate(navDir,varStart,pvarEndUpAt);
			return S_FALSE;
		}
		if (varStart.lVal<1)
			pvarEndUpAt->vt=VT_EMPTY;
		else
		{
			pvarEndUpAt->vt=VT_I4;
			pvarEndUpAt->lVal=varStart.lVal-1;
		}
		break;

		// Unsupported directions.
	case NAVDIR_LEFT:
	case NAVDIR_RIGHT:
		if (varStart.lVal==CHILDID_SELF)
		{
			if (m_pStdAccessible)
				return m_pStdAccessible->accNavigate(navDir,varStart,pvarEndUpAt);
		}
		return S_FALSE;
	}
	return S_OK;
}

HRESULT STDMETHODCALLTYPE CMenuAccessible::accHitTest( long xLeft, long yTop, VARIANT *pvarChild )
{
	if (!m_pOwner) return RPC_E_DISCONNECTED;
	POINT pt={xLeft,yTop};
	RECT rc;
	m_pOwner->GetWindowRect(&rc);
	if (!PtInRect(&rc,pt))
	{
		pvarChild->vt=VT_EMPTY;
		return S_FALSE;
	}
	for (std::vector<CWindow>::iterator it=m_pOwner->m_Toolbars.begin();it!=m_pOwner->m_Toolbars.end();++it)
	{
		POINT pt2=pt;
		it->ScreenToClient(&pt2);
		int index=(int)it->SendMessage(TB_HITTEST,0,(LPARAM)&pt2);
		if (index>=0)
		{
			TBBUTTON button;
			it->SendMessage(TB_GETBUTTON,index,(LPARAM)&button);
			pvarChild->vt=VT_I4;
			pvarChild->lVal=(int)button.dwData-CMenuContainer::ID_OFFSET+1;
			break;
		}
	}
	return S_OK;
}

HRESULT STDMETHODCALLTYPE CMenuAccessible::accDoDefaultAction( VARIANT varChild )
{
	if (!m_pOwner) return RPC_E_DISCONNECTED;
	if (varChild.vt!=VT_I4) return E_INVALIDARG;
	if (varChild.lVal==CHILDID_SELF)
	{
		// close
		for (std::vector<CMenuContainer*>::reverse_iterator it=CMenuContainer::s_Menus.rbegin();*it!=m_pOwner;++it)
			(*it)->PostMessage(WM_CLOSE);
		m_pOwner->PostMessage(WM_CLOSE);
		return S_OK;
	}
	int index=varChild.lVal-1;
	if (index<0 || index>=(int)m_pOwner->m_Items.size())
		return S_FALSE;
	// open or execute
	const CMenuContainer::MenuItem &item=m_pOwner->m_Items[index];
	if (item.id!=MENU_SEPARATOR && item.id!=MENU_EMPTY)
		m_pOwner->ActivateItem(index,item.bFolder?CMenuContainer::ACTIVATE_OPEN:CMenuContainer::ACTIVATE_EXECUTE,NULL);
	return S_OK;
}
