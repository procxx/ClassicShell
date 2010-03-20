// ## MenuContainer.h
// Classic Shell (c) 2009-2010, Ivo Beltchev
// The sources for Classic Shell are distributed under the MIT open source license

// DragDrop.cpp - handles the drag and drop functionality of CMenuContainer

#include "stdafx.h"
#include "MenuContainer.h"
#include "FNVHash.h"
#include <algorithm>

// CIDropSource - a basic IDropSource implementation. nothing to see here
class CIDropSource: public IDropSource
{
public:
	CIDropSource( bool bRight ) { m_bRight=bRight; m_bClosed=false; m_Time=GetMessageTime(); }
	// IUnknown
	virtual STDMETHODIMP QueryInterface( REFIID riid, void **ppvObject )
	{
		*ppvObject=NULL;
		if (IID_IUnknown==riid || IID_IDropSource==riid)
		{
			*ppvObject=(IDropSource*)this;
			return S_OK;
		}
		return E_NOINTERFACE;
	}

	virtual ULONG STDMETHODCALLTYPE AddRef( void ) { return 1; }
	virtual ULONG STDMETHODCALLTYPE Release( void ) { return 1; }

	// IDropSource
	virtual STDMETHODIMP QueryContinueDrag( BOOL fEscapePressed, DWORD grfKeyState )
	{
		if (!m_bClosed)
		{
			// if the mouse is outside of the menu for more than 4 seconds close the menu
			DWORD pos=GetMessagePos();
			POINT pt={(short)LOWORD(pos),(short)HIWORD(pos)};
			HWND hWnd=WindowFromPoint(pt);
			if (hWnd) hWnd=GetAncestor(hWnd,GA_ROOT);
			wchar_t name[256];
			if (hWnd)
				GetClassName(hWnd,name,_countof(name));
			else
				name[0]=0;
			if (_wcsicmp(name,L"ClassicShell.CMenuContainer")==0)
			{
				m_Time=GetMessageTime();
			}
			else
			{
				int dt=GetMessageTime()-m_Time;
				if (dt>4000)
				{
					m_bClosed=true;
					CMenuContainer::HideStartMenu();
				}
			}
		}
		if (m_bRight)
		{
			if (fEscapePressed || (grfKeyState&MK_LBUTTON))
				return DRAGDROP_S_CANCEL;
			if (!(grfKeyState&MK_RBUTTON))
				return DRAGDROP_S_DROP;
		}
		else
		{
			if (fEscapePressed || (grfKeyState&MK_RBUTTON))
				return DRAGDROP_S_CANCEL;
			if (!(grfKeyState&MK_LBUTTON))
				return DRAGDROP_S_DROP;
		}
		return S_OK;

	}

	virtual STDMETHODIMP GiveFeedback( DWORD dwEffect )
	{
		return DRAGDROP_S_USEDEFAULTCURSORS;
	}

	bool IsClosed( void ) { return m_bClosed; }

private:
	bool m_bRight;
	bool m_bClosed;
	long m_Time;
};

///////////////////////////////////////////////////////////////////////////////

bool CMenuContainer::DragOut( int index )
{
	if (!(m_Options&CONTAINER_DRAG) || s_bNoDragDrop) return false;
	const MenuItem &item=m_Items[index];
	if (!item.pItem1 || item.id!=MENU_NO) return false;

	bool bLeft=(GetKeyState(VK_LBUTTON)<0);
	bool bRight=(GetKeyState(VK_RBUTTON)<0);
	if (!bLeft && !bRight) return false;

	// get IDataObject for the current item
	CComPtr<IShellFolder> pFolder;
	PCUITEMID_CHILD pidl;
	if (FAILED(SHBindToParent(item.pItem1,IID_IShellFolder,(void**)&pFolder,&pidl)))
		return true;

	CComPtr<IDataObject> pDataObj;
	if (FAILED(pFolder->GetUIObjectOf(NULL,1,&pidl,IID_IDataObject,NULL,(void**)&pDataObj)))
		return true;

	// force synchronous operation
	CComQIPtr<IAsyncOperation> pAsync=pDataObj;
	if (pAsync)
		pAsync->SetAsyncMode(FALSE);

	// do drag drop
	s_pDragSource=this;
	m_DragIndex=index;
	CIDropSource src(!bLeft);
	DWORD dwEffect=DROPEFFECT_COPY|DROPEFFECT_MOVE|DROPEFFECT_LINK;
	HRESULT res=SHDoDragDrop(NULL,pDataObj,&src,dwEffect,&dwEffect);

	s_pDragSource=NULL;

	if (src.IsClosed())
	{
		for (std::vector<CMenuContainer*>::iterator it=s_Menus.begin();it!=s_Menus.end();++it)
			if (!(*it)->m_bDestroyed)
				(*it)->PostMessage(WM_CLOSE);
		return true;
	}

	if (res==DRAGDROP_S_DROP && !m_bDestroyed)
	{
		// check if the item still exists. refresh the menu if it doesn't
		SFGAOF flags=SFGAO_VALIDATE;
		if (FAILED(pFolder->GetAttributesOf(1,&pidl,&flags)))
		{
			SetActiveWindow();
			// close all submenus
			for (int i=(int)s_Menus.size()-1;s_Menus[i]!=this;i--)
				if (!s_Menus[i]->m_bDestroyed)
					s_Menus[i]->DestroyWindow();
			// update menu
			PostRefreshMessage();
		}
	}

	// activate the top non-destroyed menu
	for (int i=(int)s_Menus.size()-1;i>=0;i--)
		if (!s_Menus[i]->m_bDestroyed)
		{
			SetForegroundWindow(s_Menus[i]->m_hWnd);
			s_Menus[i]->SetActiveWindow();
			break;
		}

	return true;
}

HRESULT STDMETHODCALLTYPE CMenuContainer::DragEnter( IDataObject *pDataObj, DWORD grfKeyState, POINTL pt, DWORD *pdwEffect )
{
	s_bRightDrag=(grfKeyState&MK_RBUTTON)!=0;
	if (m_pDropTargetHelper)
	{
		POINT p={pt.x,pt.y};
		m_pDropTargetHelper->DragEnter(m_hWnd,pDataObj,&p,*pdwEffect);
	}
	if (!m_pParent && m_Items[0].id==MENU_EMPTY && !s_bShowTopEmpty)
	{
		// when dragging over the main menu, show an (Empty) item at the top so the user can drop items there
		s_bShowTopEmpty=true;
		InitWindow();
	}
	m_DragHoverTime=GetMessageTime()-10000;
	m_DragHoverItem=-1;
	m_pDragObject=pDataObj;
	return S_OK;
}

HRESULT STDMETHODCALLTYPE CMenuContainer::DragOver( DWORD grfKeyState, POINTL pt, DWORD *pdwEffect )
{
	s_bRightDrag=(grfKeyState&MK_RBUTTON)!=0;
	grfKeyState&=MK_SHIFT|MK_CONTROL;
	if (grfKeyState==MK_SHIFT)
		*pdwEffect&=DROPEFFECT_MOVE;
	else if (grfKeyState==MK_CONTROL)
		*pdwEffect&=DROPEFFECT_COPY;
	else if (grfKeyState==0 && s_pDragSource==this)
		*pdwEffect&=DROPEFFECT_MOVE;
	else
		*pdwEffect&=((s_pDragSource && (s_pDragSource->m_Options&CONTAINER_PROGRAMS))?DROPEFFECT_MOVE:DROPEFFECT_LINK);

	// only accept CFSTR_SHELLIDLIST data
	FORMATETC format={s_ShellFormat,NULL,DVASPECT_CONTENT,-1,TYMED_HGLOBAL};
	if (s_bNoDragDrop || !m_pDropFolder || !(m_Options&CONTAINER_DROP) || m_pDragObject->QueryGetData(&format)!=S_OK)
		*pdwEffect=DROPEFFECT_NONE;

	POINT p={pt.x,pt.y};
	if (m_pDropTargetHelper)
	{
		m_pDropTargetHelper->DragOver(&p,*pdwEffect);
	}

	ScreenToClient(&p);
	int index=HitTest(p);
	if (index>=0)
	{
		// set the new insert mark
		int mark=index;
		bool bAfter=false;
		RECT rc;
		GetItemRect(index,rc);
		int y=(rc.top+rc.bottom)/2;
		if (p.y<y)
		{
			// insert above
			if (m_Items[index].id!=MENU_NO && m_Items[index].id!=MENU_EMPTY && (index==0 || m_Items[index-1].id!=MENU_NO))
				mark=-1;
		}
		else
		{
			// insert below
			bAfter=true;
			if (m_Items[index].id!=MENU_NO && m_Items[index].id!=MENU_EMPTY && (index==m_Items.size()-1 || m_Items[index+1].id!=MENU_NO))
				mark=-1;
		}
		if (mark==-1 && m_Items[index].bFolder && (m_Items[index].bPrograms || m_Items[index].id==MENU_NO))
		{
			SetHotItem(index);
		}
		else
		{
			SetHotItem(-1);
		}
		if ((m_Options&CONTAINER_AUTOSORT) && s_pDragSource==this)
			mark=-1;
		SetInsertMark(mark,bAfter);
	}
	else
	{
		// clear the insert mark
		SetInsertMark(-1,false);
	}
	UpdateScroll(&p);

	// check if the hover delay is done and it's time to open the item
	if (index>=0 && index==m_DragHoverItem)
	{
		if ((GetMessageTime()-m_DragHoverTime)>(int)s_HoverTime && m_Submenu!=m_DragHoverItem)
		{
			// expand m_DragHoverItem
			if (!m_Items[index].bFolder || m_Items[index].pItem1)
				ActivateItem(index,ACTIVATE_OPEN,NULL);
			if (!m_Items[index].bFolder)
				SetHotItem(-1);
		}
	}
	else
	{
		m_DragHoverItem=index;
		m_DragHoverTime=GetMessageTime();
	}
	
	return S_OK;
}

HRESULT STDMETHODCALLTYPE CMenuContainer::DragLeave( void )
{
	if (m_pDropTargetHelper)
		m_pDropTargetHelper->DragLeave();
	SetInsertMark(-1,false);
	m_pDragObject.Release();
	UpdateScroll(NULL);
	return S_OK;
}

HRESULT STDMETHODCALLTYPE CMenuContainer::Drop( IDataObject *pDataObj, DWORD grfKeyState, POINTL pt, DWORD *pdwEffect )
{
	m_pDragObject.Release();

	grfKeyState&=MK_SHIFT|MK_CONTROL;
	if (!s_bRightDrag)
	{
		if (grfKeyState==MK_SHIFT)
			*pdwEffect&=DROPEFFECT_MOVE;
		else if (grfKeyState==MK_CONTROL)
			*pdwEffect&=DROPEFFECT_COPY;
		else if (grfKeyState==0 && s_pDragSource==this)
			*pdwEffect&=DROPEFFECT_MOVE;
		else
			*pdwEffect&=((s_pDragSource && (s_pDragSource->m_Options&CONTAINER_PROGRAMS))?DROPEFFECT_MOVE:DROPEFFECT_LINK);
		grfKeyState=0;
	}
	else if (!grfKeyState && (*pdwEffect&DROPEFFECT_LINK))
	{
		// when a file is dragged to the start menu he usually wants to make a shortcut
		// so when right-dragging, and linking is allowed, make it the default
		grfKeyState=MK_SHIFT|MK_CONTROL;
	}

	if (m_pDropTargetHelper)
	{
		POINT p={pt.x,pt.y};
		m_pDropTargetHelper->Drop(pDataObj,&p,*pdwEffect);
	}

	if (m_InsertMark==-1) return S_OK;

	int before=m_InsertMark;
	if (before<0) return S_OK;
	if (m_bInsertAfter && (before!=0 || m_Items[0].id!=MENU_EMPTY))
		before++;

	// clear the insert mark
	SetInsertMark(-1,false);

	if (s_pDragSource==this && (*pdwEffect&DROPEFFECT_MOVE))
	{
		if (before==m_DragIndex || before==m_DragIndex+1)
			return S_OK;
		// dropped in the same menu, just rearrange the items
		PlayMenuSound(SOUND_DROP);
		if (!(m_Options&CONTAINER_AUTOSORT))
		{
			std::vector<SortMenuItem> items;
			for (std::vector<MenuItem>::const_iterator it=m_Items.begin();it!=m_Items.end();++it)
				if (it->id==MENU_NO)
				{
					SortMenuItem item={it->name,it->nameHash,it->bFolder};
					items.push_back(item);
				}
			SortMenuItem drag=items[m_DragIndex];
			items.erase(items.begin()+m_DragIndex);
			if (before>m_DragIndex)
				before--;
			items.insert(items.begin()+before,drag);
			SaveItemOrder(items);
			PostRefreshMessage();
		}
	}
	else if (m_pDropFolder)
	{
		// simulate dropping the object into the original folder
		PlayMenuSound(SOUND_DROP);
		CComPtr<IDropTarget> pTarget;
		if (FAILED(m_pDropFolder->CreateViewObject(m_hWnd,IID_IDropTarget,(void**)&pTarget)))
			return S_OK;
		DWORD dwEffect=*pdwEffect;
		if (FAILED(pTarget->DragEnter(pDataObj,grfKeyState,pt,&dwEffect)))
			return S_OK;

		if (s_bRightDrag)
		{
			dwEffect=*pdwEffect;
			pTarget->DragOver(MK_RBUTTON|MK_SHIFT|MK_CONTROL,pt,&dwEffect);
		}
		else
			pTarget->DragOver(grfKeyState,pt,pdwEffect);
		CComQIPtr<IAsyncOperation> pAsync=pDataObj;
		if (pAsync)
			pAsync->SetAsyncMode(FALSE);
		for (std::vector<CMenuContainer*>::iterator it=s_Menus.begin();it!=s_Menus.end();++it)
			if (!(*it)->m_bDestroyed)
				(*it)->EnableWindow(FALSE); // disable all menus
		CMenuContainer *pOld=s_pDragSource;
		if (!s_pDragSource) s_pDragSource=this; // HACK: ensure s_pDragSource is not NULL even if dragging from external source (prevents the menu from closing)
		pTarget->Drop(pDataObj,grfKeyState,pt,pdwEffect);
		s_pDragSource=pOld;
		for (std::vector<CMenuContainer*>::iterator it=s_Menus.begin();it!=s_Menus.end();++it)
			if (!(*it)->m_bDestroyed)
				(*it)->EnableWindow(TRUE); // enable all menus
		SetForegroundWindow(m_hWnd);
		SetActiveWindow();
		SetFocus();

		if (!(m_Options&CONTAINER_AUTOSORT))
		{
			std::vector<SortMenuItem> items;
			for (std::vector<MenuItem>::const_iterator it=m_Items.begin();it!=m_Items.end();++it)
				if (it->id==MENU_NO)
				{
					SortMenuItem item={it->name,it->nameHash,it->bFolder};
					items.push_back(item);
				}
			SortMenuItem ins={L"",CalcFNVHash(L""),false};
			items.insert(items.begin()+before,ins);
			SaveItemOrder(items);
		}
		PostRefreshMessage();
	}
	return S_OK;
}
