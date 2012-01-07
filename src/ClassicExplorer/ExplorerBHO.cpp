// Classic Shell (c) 2009-2011, Ivo Beltchev
// The sources for Classic Shell are distributed under the MIT open source license

// ExplorerBHO.cpp : Implementation of CExplorerBHO

#include "stdafx.h"
#include "ExplorerBHO.h"
#include "Settings.h"
#include "ResourceHelper.h"
#include "resource.h"
#include "SettingsUI.h"
#include "Translations.h"
#include "dllmain.h"
#include <uxtheme.h>
#include <dwmapi.h>
#include <Ntquery.h>
#include <algorithm>

// CExplorerBHO - a browser helper object that implements Alt+Enter for the folder tree

const UINT_PTR TIMER_NAVIGATE='CLSH';

int CExplorerBHO::s_AutoNavDelay;

static PIDLIST_ABSOLUTE GetSelectedItem( HWND hwndTree )
{
	// find the PIDL of the selected item (combine all child PIDLs from the current item and its parents)
	HTREEITEM hItem=TreeView_GetSelection(hwndTree);
	PIDLIST_ABSOLUTE pidl=NULL;
	while (hItem)
	{
		TVITEMEX info={TVIF_PARAM,hItem};
		TreeView_GetItem(hwndTree,&info);
		PIDLIST_RELATIVE **pidl1=(PIDLIST_RELATIVE**)info.lParam;
		if (!pidl1 || !*pidl1 || !**pidl1)
		{
			if (pidl) ILFree(pidl);
			pidl=NULL;
			break;
		}
		PIDLIST_ABSOLUTE pidl2=pidl?ILCombine((PIDLIST_ABSOLUTE)**pidl1,pidl):(PIDLIST_ABSOLUTE)ILClone(**pidl1);
		if (pidl) ILFree(pidl);
		pidl=pidl2;
		hItem=TreeView_GetParent(hwndTree,hItem);
	}
	return pidl;
}

LRESULT CALLBACK CExplorerBHO::SubclassTreeParentProc( HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData )
{
	// when the tree selection changes start a timer to navigate to the new folder in 100ms
	if (uMsg==WM_NOTIFY && ((NMHDR*)lParam)->code==TVN_SELCHANGED)
	{
		if (GetSettingInt(L"AutoNavigate")==2 || ((NMTREEVIEW*)lParam)->action==TVC_BYKEYBOARD)
			SetTimer(((NMHDR*)lParam)->hwndFrom,TIMER_NAVIGATE,s_AutoNavDelay,NULL);
	}
	return DefSubclassProc(hWnd,uMsg,wParam,lParam);
}

// Subclass the tree control to:
//   - support Alt+Enter
//   - navigate to the new folder when you go up/down with the keyboard
//   - fix the random scrolling of the tree when a folder is expanded
//   - change the tree styles to achieve different looks
LRESULT CALLBACK CExplorerBHO::SubclassTreeProc( HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData )
{
	if (uMsg==TVM_ENSUREVISIBLE && (dwRefData&1))
	{
		// HACK! there is a bug in Win7 Explorer and when the selected folder is expanded for the first time it sends TVM_ENSUREVISIBLE for
		// the root tree item. This causes the navigation pane to scroll up. To work around the bug we ignore TVM_ENSUREVISIBLE if it tries
		// to show the root item and it is not selected
		HTREEITEM hItem=(HTREEITEM)lParam;
		if (!TreeView_GetParent(hWnd,hItem) && !(TreeView_GetItemState(hWnd,hItem,TVIS_SELECTED)&TVIS_SELECTED))
			return 0;
	}
	if (uMsg==WM_TIMER && wParam==TIMER_NAVIGATE)
	{
		// time to navigate to the selected folder (only if different from the current folder)
		KillTimer(hWnd,TIMER_NAVIGATE);
		if (GetFocus()!=hWnd)
			return 0;
		PIDLIST_ABSOLUTE pidl=GetSelectedItem(hWnd);
		if (pidl)
		{
			bool bSameFolder=false;
			CExplorerBHO *pThis=GetTlsData()->bho;
			CComPtr<IShellView> pView;
			if (pThis->m_pBrowser && SUCCEEDED(pThis->m_pBrowser->QueryActiveShellView(&pView)))
			{
				CComQIPtr<IFolderView> pView2=pView;

				CComPtr<IPersistFolder2> pFolder;
				PIDLIST_ABSOLUTE pidl2;
				if (pView2 && SUCCEEDED(pView2->GetFolder(IID_IPersistFolder2,(void**)&pFolder)) && SUCCEEDED(pFolder->GetCurFolder(&pidl2)) && pidl2)
				{
					if (ILIsEqual(pidl,pidl2))
						bSameFolder=true;
					ILFree(pidl2);
				}
			}
			ILFree(pidl);
			if (!bSameFolder)
				PostMessage(hWnd,WM_KEYDOWN,VK_SPACE,0);
		}
		return 0;
	}

	// ignore the Space character and Alt+Enter syscharacter (to stop the tree view from beeping)
	if (uMsg==WM_CHAR && wParam==' ')
		return 0;
	if (uMsg==WM_SYSCHAR && wParam==VK_RETURN)
		return 0;

	if (uMsg==WM_SYSKEYDOWN && wParam==VK_RETURN)
	{
		// Alt+Enter is pressed
		// if this message was for the folder tree, show the properties of the selected item
		if (GetSettingBool(L"AltEnter") && ShowTreeProperties(hWnd))
			return 0;
	}

	if (uMsg==TVM_SETEXTENDEDSTYLE && wParam==(TVS_EX_FADEINOUTEXPANDOS|TVS_EX_AUTOHSCROLL|0x80000000) && lParam==0)
	{
		wParam&=0x7FFFFFFF;

		if (GetSettingInt(L"AutoNavigate")>0)
			SetWindowSubclass(GetParent(hWnd),SubclassTreeParentProc,'CLSH',0);

		if (!GetSettingBool(L"NoFadeButtons"))
			wParam&=~TVS_EX_FADEINOUTEXPANDOS;

		int indent=-1;
		if (GetSettingBool(L"FullIndent"))
			indent=0;

		int treeStyle=GetSettingInt(L"TreeStyle");
		DWORD style=GetWindowLong(hWnd,GWL_STYLE);
		if (treeStyle!=STYLE_VISTA)
		{
			SetWindowTheme(hWnd,NULL,NULL);
			if (treeStyle==STYLE_SIMPLE)
			{
				style|=TVS_SINGLEEXPAND|TVS_TRACKSELECT;
				style&=~TVS_HASLINES;
			}
			else
			{
				style|=TVS_HASLINES;
				style&=~(TVS_SINGLEEXPAND|TVS_TRACKSELECT);
				wParam|=TVS_EX_FADEINOUTEXPANDOS;
				HIMAGELIST images=TreeView_GetImageList(hWnd,TVSIL_NORMAL);
				int cx, cy;
				ImageList_GetIconSize(images,&cx,&cy);
				indent=cx+3;
			}
		}
		int scroll=GetSettingInt(L"HScrollbar");

		if ((scroll==0 && treeStyle==STYLE_VISTA) || scroll==1)
			wParam&=~TVS_EX_AUTOHSCROLL;
		if ((scroll==0 && treeStyle!=STYLE_VISTA) || scroll==2)
			style&=~TVS_NOHSCROLL;

		SetWindowLong(hWnd,GWL_STYLE,style);

		if (indent>=0)
			TreeView_SetIndent(hWnd,indent);

		int d=GetSettingInt(L"TreeItemSpacing");
		if (d)
			TreeView_SetItemHeight(hWnd,TreeView_GetItemHeight(hWnd)+d);

		if (wParam==0)
			return 0;
	}

	if (uMsg==WM_SETTINGCHANGE)
	{
		LRESULT res=DefSubclassProc(hWnd,uMsg,wParam,lParam);
		int indent=-1;
		if (GetSettingBool(L"FullIndent"))
			indent=0;

		if (GetSettingInt(L"TreeStyle")==STYLE_CLASSIC)
		{
			HIMAGELIST images=TreeView_GetImageList(hWnd,TVSIL_NORMAL);
			int cx, cy;
			ImageList_GetIconSize(images,&cx,&cy);
			indent=cx+3;
		}
		if (indent>=0)
			TreeView_SetIndent(hWnd,indent);

		int d=GetSettingInt(L"TreeItemSpacing");
		if (d)
			TreeView_SetItemHeight(hWnd,TreeView_GetItemHeight(hWnd)+d);
		return res;
	}

	return DefSubclassProc(hWnd,uMsg,wParam,lParam);
}

LRESULT CALLBACK CExplorerBHO::HookExplorer( int nCode, WPARAM wParam, LPARAM lParam )
{
	// wait for the tree control to be created and subclass it
	if (nCode==HCBT_CREATEWND)
	{
		HWND hWnd=(HWND)wParam;
		CBT_CREATEWND *create=(CBT_CREATEWND*)lParam;
		if (create->lpcs->lpszClass>(LPTSTR)0xFFFF && _wcsicmp(create->lpcs->lpszClass,WC_TREEVIEW)==0)
		{
			HWND parent=GetAncestor(create->lpcs->hwndParent,GA_ROOT);
			wchar_t name[256];
			if (GetClassName(parent,name,_countof(name)) && _wcsicmp(name,L"CabinetWClass")==0)
			{
				DWORD_PTR settings=0;
				DWORD version=LOWORD(GetVersion());
				if (version!=0x0006 && GetSettingBool(L"FixFolderScroll"))
					settings|=1;
				SetWindowSubclass(hWnd,SubclassTreeProc,'CLSH',settings);
				PostMessage(hWnd,TVM_SETEXTENDEDSTYLE,TVS_EX_FADEINOUTEXPANDOS|TVS_EX_AUTOHSCROLL|0x80000000,0);
				if (version==0x0006)
				{
					// on Vista we can unhook right now. on Win7 we keep the hook because sometimes the tree control can get destroyed and recreated
					TlsData *pTlsData=GetTlsData();
					UnhookWindowsHookEx(pTlsData->bho->m_Hook);
					pTlsData->bho->m_Hook=NULL;
				}
				return 0;
			}
		}
	}
	return CallNextHookEx(NULL,nCode,wParam,lParam);
}

LRESULT CALLBACK CExplorerBHO::HookKeyboard( int nCode, WPARAM wParam, LPARAM lParam )
{
	// wait for the tree control to be created and subclass it
	if (nCode==HC_ACTION)
	{
		TlsData *pTlsData=GetTlsData();
		if (wParam==pTlsData->bho->m_AltD && (lParam&0x20000000))
		{
			if (lParam&0x80000000)
			{
				return 0;
			}
			else
			{
				if (IsWindow(pTlsData->bho->m_Breadcrumbs) && IsWindowVisible(pTlsData->bho->m_Breadcrumbs))
				{
					SetFocus(pTlsData->bho->m_Breadcrumbs);
					SendMessage(pTlsData->bho->m_Breadcrumbs,WM_KEYDOWN,VK_SPACE,0);
					SendMessage(pTlsData->bho->m_Breadcrumbs,WM_KEYUP,VK_SPACE,0);
					return 0;
				}
			}
		}
		if (wParam==(pTlsData->bho->m_UpHotkey&255) && !(lParam&0x80000000))
		{
			// Backspace goes to the parent folder, but only if no window has the caret
			GUITHREADINFO info={sizeof(info)};
			if (GetGUIThreadInfo(GetCurrentThreadId(),&info) && !info.hwndCaret)
			{
				bool bShift1=(pTlsData->bho->m_UpHotkey&(HOTKEYF_SHIFT<<8))!=0;
				bool bCtrl1=(pTlsData->bho->m_UpHotkey&(HOTKEYF_CONTROL<<8))!=0;
				bool bAlt1=(pTlsData->bho->m_UpHotkey&(HOTKEYF_ALT<<8))!=0;
				bool bShift2=GetKeyState(VK_SHIFT)<0;
				bool bCtrl2=GetKeyState(VK_CONTROL)<0;
				bool bAlt2=GetKeyState(VK_MENU)<0;
				if (bShift1==bShift2 && bCtrl1==bCtrl2 && bAlt1==bAlt2)
				{
					pTlsData->bho->m_pBrowser->BrowseObject(NULL,SBSP_SAMEBROWSER|SBSP_PARENT);
					return 1;
				}
			}
		}
	}
	return CallNextHookEx(NULL,nCode,wParam,lParam);
}

// Subclass the statusbar to:
//   - show free disk space
//   - show the total size of the selected files
LRESULT CALLBACK CExplorerBHO::SubclassStatusProc( HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData )
{
	wchar_t buf[1024];
	if (uMsg==WM_PAINT && ((CExplorerBHO*)uIdSubclass)->m_bForceRefresh)
	{
		// sometimes Explorer doesn't fully initialize the status bar on Windows 7 and leaves it with 1 part
		// in such case force the view to refresh after the status bar is fully visible
		((CExplorerBHO*)uIdSubclass)->m_bForceRefresh=false;
		if (SendMessage(hWnd,SB_GETPARTS,0,0)<=1)
			PostMessage(GetParent(hWnd),WM_COMMAND,41504,0); // Refresh command
	}
	if (uMsg==SB_SETTEXT && LOWORD(wParam)==0)
	{
		// when the text of the first part is changing

		// recalculate the selection size on a timer. this way if the status text is changed frequently
		// the recalculation will not happen every time
		SetTimer(hWnd,uIdSubclass,10,NULL);

		if (dwRefData&SPACE_WIN7)
		{
			if (((CExplorerBHO*)uIdSubclass)->m_bResetStatus && SendMessage(hWnd,SB_GETPARTS,0,0)<=1)
			{
				// HACK! there is a bug in Win7 and when the Explorer window is created it doesn't correctly
				// initialize the status bar to have 3 parts. as soon as the user resizes the window the
				// 3 parts appear. so here we resize the parent of the status bar to create the 3 parts.
				HWND parent=GetParent(hWnd);
				RECT rc;
				GetWindowRect(parent,&rc);
				SetWindowPos(parent,NULL,0,0,rc.right-rc.left+1,rc.bottom-rc.top,SWP_NOZORDER|SWP_NOMOVE);
				SetWindowPos(parent,NULL,0,0,rc.right-rc.left,rc.bottom-rc.top,SWP_NOZORDER|SWP_NOMOVE);
				// the first time the status text is set it is too early. so we do this until we get at lest 2 parts
				if (SendMessage(hWnd,SB_GETPARTS,0,0)>1)
					((CExplorerBHO*)uIdSubclass)->m_bResetStatus=false;
			}
		}
		IShellBrowser *pBrowser=((CExplorerBHO*)uIdSubclass)->m_pBrowser;
		CComPtr<IShellView> pView;
		if (pBrowser && SUCCEEDED(pBrowser->QueryActiveShellView(&pView)))
		{
			CComQIPtr<IFolderView> pView2=pView;

			CComPtr<IPersistFolder2> pFolder;
			if (pView2 && SUCCEEDED(pView2->GetFolder(IID_IPersistFolder2,(void**)&pFolder)))
			{
				bool bInfoTip=false;
				int count;
				if ((dwRefData&SPACE_INFOTIP) && SUCCEEDED(pView2->ItemCount(SVGIO_SELECTION,&count)) && count==1)
				{
					// if only one item is selected, show its info in the status bar
					CComPtr<IEnumIDList> pEnum;
					PITEMID_CHILD child;
					if (SUCCEEDED(pView2->Items(SVGIO_SELECTION,IID_IEnumIDList,(void**)&pEnum)) && pEnum && pEnum->Next(1,&child,NULL)==S_OK)
					{
						CComQIPtr<IShellFolder> pFolder2=pFolder;
						if (pFolder2)
						{
							CComPtr<IQueryInfo> pQueryInfo;
							if (SUCCEEDED(pFolder2->GetUIObjectOf(NULL,1,(PCUITEMID_CHILD*)&child,IID_IQueryInfo,NULL,(void**)&pQueryInfo)))
							{
								wchar_t *tip;
								if (SUCCEEDED(pQueryInfo->GetInfoTip(QITIPF_DEFAULT|QITIPF_SINGLELINE,&tip)) && tip)
								{
									Strcpy(buf,_countof(buf),tip);
									for (wchar_t *p=buf;*p;p++)
										if (*p=='\t')
											*p=' ';
									CoTaskMemFree(tip);
									bInfoTip=true;
									lParam=(LPARAM)buf;
								}
							}
						}
					}
				}

				if (!bInfoTip && (dwRefData&SPACE_WIN7))
				{
					// show the free space of the drive containing the current folder
					// also works for network locations
					PIDLIST_ABSOLUTE pidl;
					if (SUCCEEDED(pFolder->GetCurFolder(&pidl)))
					{
						if (SHGetPathFromIDList(pidl,buf))
						{
							ULARGE_INTEGER size;
							if (GetDiskFreeSpaceEx(buf,NULL,NULL,&size))
							{
								const wchar_t *text=(wchar_t*)lParam;
								wchar_t str[100];
								StrFormatByteSize64(size.QuadPart,str,_countof(str));
								Sprintf(buf,_countof(buf),FindTranslation(L"Status.FreeSpace",L"%s (Disk free space: %s)"),text,str);
								lParam=(LPARAM)buf;
							}
						}
						ILFree(pidl);
					}
				}
			}
		}
	}
	if (uMsg==SB_SETTEXT && LOWORD(wParam)==1)
	{
		return 0;
	}

	if (uMsg==WM_TIMER && wParam==uIdSubclass)
	{
		// recalculate the total size of the selected files and show it in part 2 of the status bar
		KillTimer(hWnd,wParam);
		IShellBrowser *pBrowser=((CExplorerBHO*)uIdSubclass)->m_pBrowser;
		__int64 size=-1;
		CComPtr<IShellView> pView;
		if (pBrowser && SUCCEEDED(pBrowser->QueryActiveShellView(&pView)))
		{
			CComQIPtr<IFolderView> pView2=pView;
			CComPtr<IPersistFolder2> pFolder;
			PIDLIST_ABSOLUTE pidl;
			if (pView2 && SUCCEEDED(pView2->GetFolder(IID_IPersistFolder2,(void**)&pFolder)) && SUCCEEDED(pFolder->GetCurFolder(&pidl)))
			{
				CComQIPtr<IShellFolder2> pFolder2=pFolder;
				UINT type=SVGIO_SELECTION;
				int count;
				if ((FAILED(pView2->ItemCount(SVGIO_SELECTION,&count)) || count==0))
					type=SVGIO_ALLVIEW;
				CComPtr<IEnumIDList> pEnum;
				if (SUCCEEDED(pView2->Items(type,IID_IEnumIDList,(void**)&pEnum)) && pEnum)
				{
					PITEMID_CHILD child;
					SHCOLUMNID column={PSGUID_STORAGE,PID_STG_SIZE};
					while (pEnum->Next(1,&child,NULL)==S_OK)
					{
						CComVariant var;
						if (SUCCEEDED(pFolder2->GetDetailsEx(child,&column,&var)) && var.vt==VT_UI8)
						{
							if (size<0)
								size=var.ullVal;
							else
								size+=var.ullVal;
						}
						ILFree(child);
					}
				}
				ILFree(pidl);
			}
		}
		if (size>=0)
		{
			// format the file size as KB, MB, etc
			StrFormatByteSize64(size,buf,_countof(buf));
		}
		else
			buf[0]=0;
		DefSubclassProc(hWnd,SB_SETTEXT,1,(LPARAM)buf);
	}
	return DefSubclassProc(hWnd,uMsg,wParam,lParam);
}

// Subclass the rebar in the title bar to handle the title bar Up button
LRESULT CALLBACK CExplorerBHO::RebarSubclassProc( HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData )
{
	if (uMsg==WM_NOTIFY && ((NMHDR*)lParam)->hwndFrom==(HWND)dwRefData && ((NMHDR*)lParam)->code==NM_CUSTOMDRAW)
	{
		// custom-draw the toolbar. just draw the correct icon and nothing else
		NMTBCUSTOMDRAW *pDraw=(NMTBCUSTOMDRAW*)lParam;
		if (pDraw->nmcd.dwDrawStage==CDDS_PREPAINT)
			return CDRF_NOTIFYITEMDRAW;
		if (pDraw->nmcd.dwDrawStage==CDDS_ITEMPREPAINT)
		{
			CExplorerBHO *pThis=(CExplorerBHO*)uIdSubclass;
			BOOL comp;
			if (SUCCEEDED(DwmIsCompositionEnabled(&comp)) && comp)
				FillRect(pDraw->nmcd.hdc,&pDraw->nmcd.rc,(HBRUSH)GetStockObject(BLACK_BRUSH));
			if (pDraw->nmcd.uItemState&CDIS_DISABLED)
			{
				if (pThis->m_IconDisabled)
					DrawIconEx(pDraw->nmcd.hdc,0,0,pThis->m_IconDisabled,0,0,0,NULL,DI_NORMAL|DI_NOMIRROR);
				else
					DrawIconEx(pDraw->nmcd.hdc,0,0,pThis->m_IconNormal,0,0,0,NULL,DI_NORMAL|DI_NOMIRROR);
			}
			else if (pDraw->nmcd.uItemState&CDIS_SELECTED)
			{
				if (pThis->m_IconPressed)
					DrawIconEx(pDraw->nmcd.hdc,0,0,pThis->m_IconPressed,0,0,0,NULL,DI_NORMAL|DI_NOMIRROR);
				else
					DrawIconEx(pDraw->nmcd.hdc,1,1,pThis->m_IconNormal,0,0,0,NULL,DI_NORMAL|DI_NOMIRROR);
			}
			else if (pDraw->nmcd.uItemState&CDIS_HOT)
			{
				if (pThis->m_IconHot)
					DrawIconEx(pDraw->nmcd.hdc,0,0,pThis->m_IconHot,0,0,0,NULL,DI_NORMAL|DI_NOMIRROR);
				else
					DrawIconEx(pDraw->nmcd.hdc,0,0,pThis->m_IconNormal,0,0,0,NULL,DI_NORMAL|DI_NOMIRROR);
			}
			else
				DrawIconEx(pDraw->nmcd.hdc,0,0,pThis->m_IconNormal,0,0,0,NULL,DI_NORMAL|DI_NOMIRROR);
			return CDRF_SKIPDEFAULT;
		}
	}

	if (uMsg==WM_THEMECHANGED)
	{
		// the button size is reset when the theme changes. force the correct size again
		HWND toolbar=(HWND)dwRefData;
		RECT rc;
		GetClientRect(toolbar,&rc);
		PostMessage(toolbar,TB_SETBUTTONSIZE,0,MAKELONG(rc.right,rc.bottom));
	}

	if (uMsg==WM_NOTIFY && ((NMHDR*)lParam)->hwndFrom==(HWND)dwRefData && ((NMHDR*)lParam)->code==TBN_GETINFOTIP)
	{
		// show the tip for the up button
		NMTBGETINFOTIP *pTip=(NMTBGETINFOTIP*)lParam;
		Strcpy(pTip->pszText,pTip->cchTextMax,FindTranslation(L"Toolbar.GoUp",L"Up One Level"));
		return 0;
	}

	if (uMsg==WM_NOTIFY && ((NMHDR*)lParam)->hwndFrom==(HWND)dwRefData && ((NMHDR*)lParam)->code==NM_RCLICK)
	{
		NMMOUSE *pInfo=(NMMOUSE*)lParam;
		POINT pt=pInfo->pt;
		ClientToScreen(pInfo->hdr.hwndFrom,&pt);
		ShowSettingsMenu(hWnd,pt.x,pt.y);
		return TRUE;
	}

	if (uMsg==WM_COMMAND && wParam==1)
	{
		UINT flags=(GetKeyState(VK_CONTROL)<0?SBSP_NEWBROWSER:SBSP_SAMEBROWSER);
		((CExplorerBHO*)uIdSubclass)->m_pBrowser->BrowseObject(NULL,flags|SBSP_PARENT);
	}

	CExplorerBHO *pThis=(CExplorerBHO*)uIdSubclass;
	if (pThis->m_bRemapBands)
	{
		// HACK! Explorer doesn't use RB_IDTOINDEX every time it needs to access a particular band. Since we insert the Up button in the second
		// position, the rest of the bands get offset and comedy ensues (the search box is not sized properly).
		// To fix the issue, we renumber the bands so that from the outside ours appears to be last.
		static int remapNewOld[2][4]={{3,0,1,2},{0,3,1,2}};
		static int remapOldNew[2][4]={{1,2,3,0},{0,2,3,1}};

		if (uMsg==RB_IDTOINDEX || uMsg==RB_HITTEST)
		{
			// remap the result from RB_IDTOINDEX and RB_HITTEST
			LRESULT res=DefSubclassProc(hWnd,uMsg,wParam,lParam);
			if (res<0 || res>3) return res;
			res=remapNewOld[pThis->m_UpButtonIndex-1][res];
			if (lParam && uMsg==RB_HITTEST) ((RBHITTESTINFO*)lParam)->iBand=(int)res;
			return res;
		}
		if (uMsg==RB_GETBANDBORDERS || uMsg==RB_GETBANDINFO || uMsg==RB_GETRECT || uMsg==RB_SETBANDINFO || uMsg==RB_SETBANDWIDTH)
		{
			// remap wParam for all GET/SET messages
			if (wParam>=0 && wParam<=3)
				wParam=remapOldNew[pThis->m_UpButtonIndex-1][wParam];
		}
	}

	return DefSubclassProc(hWnd,uMsg,wParam,lParam);
}

// Subclass the breadcrumbs to make them show the full path
LRESULT CALLBACK CExplorerBHO::BreadcrumbSubclassProc( HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData )
{
	CExplorerBHO *pThis=(CExplorerBHO*)uIdSubclass;
	if (*pThis->m_CurPath)
	{
		if (uMsg==WM_SETFOCUS)
		{
			if (wParam)
			{
				// see if the focus comes from the combo box. if so, most likely Escape was pressed, so just focus the main frame
				HWND from=(HWND)wParam;
				HWND combo=FindChildWindow(GetParent(GetParent(hWnd)),WC_COMBOBOXEX);
				if (combo && (combo==from || IsChild(combo,from)))
				{
					SetFocus(GetAncestor(hWnd,GA_ROOT));
					return 0;
				}
			}
			// when the breadcrumbs are focused, switch to the combobox by simulating a mouse click
			RECT rc;
			GetClientRect(hWnd,&rc);
			LPARAM pos=MAKELONG(rc.right-1,rc.bottom/2);
			DefSubclassProc(hWnd,WM_LBUTTONDOWN,MK_LBUTTON,pos);
			DefSubclassProc(hWnd,WM_LBUTTONUP,0,pos);
			return 0;
		}
		if (uMsg==WM_LBUTTONDOWN || uMsg==WM_LBUTTONDBLCLK || uMsg==WM_LBUTTONUP)
		{
			// unless the mouse is clicked on the icon, replace the mouse position with a point on the far right.
			// this will cause Explorer to switch to the combobox even when a breadcrumb is clicked
			int iconSize=GetSystemMetrics(SM_CXSMICON);
			if (!pThis->m_CurIcon || (short)LOWORD(lParam)>iconSize+3)
			{
				RECT rc;
				GetClientRect(hWnd,&rc);
				lParam=MAKELONG(rc.right-1,rc.bottom/2);
			}
		}

		if (uMsg==WM_PAINT)
		{
			// make the breadcrumbs control draw the full path like the XP address bar
			RECT rc;
			GetClientRect(hWnd,&rc);

			PAINTSTRUCT ps;
			HDC hdc=BeginPaint(hWnd,&ps);

			// we need to use buffered painting because DrawThemeTextEx with DTT_COMPOSITED requires it
			// on Vista DTT_COMPOSITED is required so that the black text doesn't get transparent. On Windows 7 regular DrawText seems to work fine
			BP_PAINTPARAMS paintParams={sizeof(paintParams)};
			paintParams.dwFlags=BPPF_ERASE;
			HDC hdcPaint=NULL;
			HPAINTBUFFER hBufferedPaint=BeginBufferedPaint(hdc,&rc,BPBF_TOPDOWNDIB,&paintParams,&hdcPaint);
			if (hdcPaint)
			{
				rc.top++;
				SendMessage(GetParent(GetParent(hWnd)),WM_PRINTCLIENT,(WPARAM)hdcPaint,PRF_CLIENT);

				// draw icon
				int iconSize=GetSystemMetrics(SM_CXSMICON);
				if (pThis->m_CurIcon)
					DrawIconEx(hdcPaint,rc.left+3,(rc.top+rc.bottom-iconSize)/2,pThis->m_CurIcon,iconSize,iconSize,0,NULL,DI_NORMAL);
				rc.left+=iconSize+8; // Not a good idea to hard-code number of pixels, but seems to work fine for different DPI settings

				// draw path
				HFONT font=(HFONT)SendMessage(hWnd,WM_GETFONT,0,0);
				HFONT font0=(HFONT)SelectObject(hdcPaint,font);
				SetBkMode(hdcPaint,TRANSPARENT);
				SetTextColor(hdcPaint,GetSysColor(COLOR_WINDOWTEXT));
				HTHEME theme=GetWindowTheme(hWnd);
				BOOL dwm;
				if (theme && SUCCEEDED(DwmIsCompositionEnabled(&dwm)) && dwm)
				{
					DTTOPTS opts={sizeof(opts),DTT_COMPOSITED|DTT_TEXTCOLOR};
					opts.crText=GetSysColor(COLOR_WINDOWTEXT);
					DrawThemeTextEx(theme,hdcPaint,0,0,pThis->m_CurPath,-1,DT_NOPREFIX|DT_VCENTER|DT_SINGLELINE,&rc,&opts);
				}
				else
				{
					DrawText(hdcPaint,pThis->m_CurPath,-1,&rc,DT_NOPREFIX|DT_VCENTER|DT_SINGLELINE);
				}
				SelectObject(hdcPaint,font0);
				EndBufferedPaint(hBufferedPaint,TRUE);
			}
			EndPaint(hWnd,&ps);
			return 0;
		}
	}
	return DefSubclassProc(hWnd,uMsg,wParam,lParam);
}

// Subclass the progress bar behind the address bar to remove the history and replace it with a list of parent folders
LRESULT CALLBACK CExplorerBHO::ProgressSubclassProc( HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData )
{
	CExplorerBHO *pThis=(CExplorerBHO*)uIdSubclass;

	if (uMsg==WM_PARENTNOTIFY && LOWORD(wParam)==WM_CREATE && !pThis->m_ComboBox.m_hWnd)
	{
		// on Windows 7 the combobox is not created at startup. so listen for child windows being created and update pThis->m_ComboBox
		HWND combo=(HWND)lParam;
		wchar_t className[256];
		GetClassName(combo,className,_countof(className));
		if (_wcsicmp(className,WC_COMBOBOXEX)==0)
			pThis->m_ComboBox=combo;
	}

	if (uMsg==WM_COMMAND && (HWND)lParam==pThis->m_ComboBox.m_hWnd)
	{
		if (HIWORD(wParam)==CBN_DROPDOWN)
		{
			// on drop down refresh the list
			pThis->ClearComboItems();

			CComPtr<IShellFolder> pDesktop;
			SHGetDesktopFolder(&pDesktop);

			{
				// add desktop
				ITEMIDLIST shEmpty={{0}};
				wchar_t *pName;
				if (SUCCEEDED(SHGetNameFromIDList((PIDLIST_ABSOLUTE)&shEmpty,SIGDN_DESKTOPABSOLUTEEDITING,&pName)))
				{
					ComboItem item={ILCloneFull((PIDLIST_ABSOLUTE)&shEmpty),0,CString(pName)};
					pThis->m_ComboItems.push_back(item);
					CoTaskMemFree(pName);
				}
			}

			if (dwRefData==2)
			{
				// enumerate all desktop items
				CComPtr<IEnumIDList> pEnum;
				if (SUCCEEDED(pDesktop->EnumObjects(NULL,SHCONTF_FOLDERS,&pEnum)) && pEnum)
				{
					PITEMID_CHILD child;
					while (pEnum->Next(1,&child,NULL)==S_OK)
					{
						STRRET str;
						if (SUCCEEDED(pDesktop->GetDisplayNameOf(child,SHGDN_INFOLDER|SHGDN_NORMAL,&str)))
						{
							wchar_t *pName;
							StrRetToStr(&str,child,&pName);
							ComboItem item={(PIDLIST_ABSOLUTE)child,1,CString(pName)};
							item.sortName=item.name;
							pThis->m_ComboItems.push_back(item);
							CoTaskMemFree(pName);
						}
						else
							ILFree(child);
					}
				}

				// sort desktop items
				std::sort(pThis->m_ComboItems.begin()+1,pThis->m_ComboItems.end());
			}

			if (dwRefData==2)
			{
				// enumerate all computer items
				PIDLIST_ABSOLUTE pidlComp;
				SHGetKnownFolderIDList(FOLDERID_ComputerFolder,0,NULL,&pidlComp);
				int index=1;
				for (int i=1;i<(int)pThis->m_ComboItems.size();i++)
					if (ILIsEqual(pidlComp,pThis->m_ComboItems[i].pidl))
					{
						index=i+1;
						break;
					}
					int index0=index;
					CComPtr<IShellFolder> pComputer;
					pDesktop->BindToObject(pidlComp,NULL,IID_IShellFolder,(void**)&pComputer);

					CComPtr<IEnumIDList> pEnum;
					if (pComputer && SUCCEEDED(pComputer->EnumObjects(NULL,SHCONTF_FOLDERS,&pEnum)) && pEnum)
					{
						PITEMID_CHILD child;
						while (pEnum->Next(1,&child,NULL)==S_OK)
						{
							STRRET str;
							if (SUCCEEDED(pComputer->GetDisplayNameOf(child,SHGDN_INFOLDER|SHGDN_NORMAL,&str)))
							{
								wchar_t *pName;
								StrRetToStr(&str,child,&pName);
								ComboItem item={ILCombine(pidlComp,child),2,CString(pName)};
								pThis->m_ComboItems.insert(pThis->m_ComboItems.begin()+index,1,item);
								CoTaskMemFree(pName);
								index++;
							}
							ILFree(child);
						}
					}
					ILFree(pidlComp);

					//sort computer items
					std::sort(pThis->m_ComboItems.begin()+index0,pThis->m_ComboItems.begin()+index);
			}

			if (pThis->m_CurPidl)
			{
				// enumerate all parent items
				PIDLIST_ABSOLUTE pidl=ILCloneFull(pThis->m_CurPidl);
				PIDLIST_ABSOLUTE pidlStart=pidl;
				int index=0;
				for (int i=0;i<(int)pThis->m_ComboItems.size();i++)
				{
					PIDLIST_ABSOLUTE p=(PIDLIST_ABSOLUTE)ILFindChild(pThis->m_ComboItems[i].pidl,pidl);
					if (p)
					{
						index=i;
						pidlStart=p;
					}
				}
				int n=0;
				for (PUIDLIST_RELATIVE child=pidl;!ILIsEmpty(child);child=ILGetNext(child))
					n++;

				int start=n;
				for (PUIDLIST_RELATIVE child=pidlStart;!ILIsEmpty(child);child=ILGetNext(child))
					start--;

				for (int i=start;i<n;i++)
				{
					wchar_t *pName=NULL;
					CComPtr<IShellFolder> pFolder;
					PCITEMID_CHILD child;
					if (SUCCEEDED(SHBindToParent(pidl,IID_IShellFolder,(void**)&pFolder,&child)))
					{
						STRRET str;
						if (SUCCEEDED(pFolder->GetDisplayNameOf(child,SHGDN_INFOLDER|SHGDN_NORMAL,&str)))
						{
							wchar_t *pName;
							StrRetToStr(&str,child,&pName);
							ComboItem item={ILCloneFull(pidl),n-(i-start),CString(pName)};
							pThis->m_ComboItems.insert(pThis->m_ComboItems.begin()+index+1,1,item);
							CoTaskMemFree(pName);
						}
					}
					ILRemoveLastID(pidl);
				}
				ILFree(pidl);
			}

			// add all sorted items to the combobox
			COMBOBOXEXITEM item={CBEIF_TEXT|CBEIF_IMAGE|CBEIF_SELECTEDIMAGE|CBEIF_INDENT|CBEIF_LPARAM};
			item.iItem=-1;
			for (std::vector<ComboItem>::const_iterator it=pThis->m_ComboItems.begin();it!=pThis->m_ComboItems.end();++it)
			{
				item.iImage=item.iSelectedImage=-1;
				item.pszText=(LPWSTR)(LPCWSTR)it->name;
				item.iIndent=it->indent;
				item.lParam=(LPARAM)it->pidl;
				SHFILEINFO info;
				if (SHGetFileInfo((LPCTSTR)it->pidl,0,&info,sizeof(info),SHGFI_PIDL|SHGFI_SYSICONINDEX|SHGFI_SMALLICON))
					item.iImage=item.iSelectedImage=info.iIcon;
				int idx=(int)pThis->m_ComboBox.SendMessage(CBEM_INSERTITEM,'CLSH',(LPARAM)&item);
				if (pThis->m_CurPidl && ILIsEqual(it->pidl,pThis->m_CurPidl))
					pThis->m_ComboBox.SendMessage(CB_SETCURSEL,idx);
			}

			return 0;
		}
		if (HIWORD(wParam)==CBN_CLOSEUP || HIWORD(wParam)==CBN_SELENDCANCEL)
		{
			// on close clear the list
			if (!pThis->m_ComboItems.empty())
			{
				pThis->m_ComboBox.SetWindowText(pThis->m_CurPath);
				pThis->ClearComboItems();
			}
			if (HIWORD(wParam)==CBN_SELENDCANCEL)
				PostMessage(hWnd,pThis->m_NavigateMsg,1,0);
			return 0;
		}
		if (HIWORD(wParam)==CBN_SELENDOK)
		{
			// when an item is selected, go to it and clear the list (selection with mouse)
			int index=(int)pThis->m_ComboBox.SendMessage(CB_GETCURSEL);
			if (index>=0)
			{
				COMBOBOXEXITEM item={CBEIF_LPARAM,index};
				pThis->m_ComboBox.SendMessage(CBEM_GETITEM,0,(LPARAM)&item);
				pThis->m_NavigatePidl=ILCloneFull((PIDLIST_ABSOLUTE)item.lParam);
				PostMessage(hWnd,pThis->m_NavigateMsg,0,0);
			}
			pThis->ClearComboItems();
			return 0;
		}
	}

	if (uMsg==WM_NOTIFY && ((NMHDR*)lParam)->code==CBEN_BEGINEDIT)
	{
		// ignore this, so Explorer doesn't add its history to the list
		return 0;
	}

	if (uMsg==WM_NOTIFY && ((NMHDR*)lParam)->code==CBEN_ENDEDIT)
	{
		if (pThis->m_ComboBox.SendMessage(CB_GETDROPPEDSTATE))
		{
			// when an item is selected, go to it and clear the list (selection with keyboard)
			NMCBEENDEDIT *pEdit=(NMCBEENDEDIT*)lParam;
			int index=pEdit->iNewSelection;
			if (index>=0)
			{
				COMBOBOXEXITEM item={CBEIF_LPARAM,index};
				pThis->m_ComboBox.SendMessage(CBEM_GETITEM,0,(LPARAM)&item);
				pThis->m_NavigatePidl=ILCloneFull((PIDLIST_ABSOLUTE)item.lParam);
				pThis->ClearComboItems();
				PostMessage(hWnd,pThis->m_NavigateMsg,0,0);
				pThis->m_ComboBox.SendMessage(CB_SHOWDROPDOWN,FALSE);
			}
			return 0;
		}
	}

	if (uMsg==pThis->m_NavigateMsg)
	{
		// navigate to the selected item
		if (wParam==0 && pThis->m_NavigatePidl && !(pThis->m_CurPidl && ILIsEqual(pThis->m_NavigatePidl,pThis->m_CurPidl)))
			pThis->m_pBrowser->BrowseObject(pThis->m_NavigatePidl,SBSP_SAMEBROWSER|SBSP_ABSOLUTE);
		else
			SetFocus(hWnd);
		if (pThis->m_NavigatePidl) ILFree(pThis->m_NavigatePidl);
		pThis->m_NavigatePidl=NULL;
	}

	return DefSubclassProc(hWnd,uMsg,wParam,lParam);
}

HRESULT STDMETHODCALLTYPE CExplorerBHO::SetSite( IUnknown *pUnkSite )
{
	IObjectWithSiteImpl<CExplorerBHO>::SetSite(pUnkSite);

	if (pUnkSite)
	{
		// hook
		GetTlsData()->bho=this;
		if (!m_Hook)
		{
			m_Hook=SetWindowsHookEx(WH_CBT,HookExplorer,NULL,GetCurrentThreadId());
		}
		CComQIPtr<IServiceProvider> pProvider=pUnkSite;

		if (pProvider)
		{
			pProvider->QueryService(SID_SShellBrowser,IID_IShellBrowser,(void**)&m_pBrowser);

			// listen for web browser notifications. we only care about DISPID_NAVIGATECOMPLETE2 and DISPID_ONQUIT
			pProvider->QueryService(SID_SWebBrowserApp,IID_IWebBrowser2,(void**)&m_pWebBrowser);
			if (m_pWebBrowser)
			{
				if (m_dwEventCookie==0xFEFEFEFE) // ATL's event cookie is 0xFEFEFEFE when the sink is not advised
					DispEventAdvise(m_pWebBrowser,&DIID_DWebBrowserEvents2);
				CRegKey regKey;
				if (regKey.Open(HKEY_CURRENT_USER,L"Software\\IvoSoft\\ClassicExplorer")==ERROR_SUCCESS)
				{
					DWORD val;
					if (regKey.QueryDWORDValue(L"ShowedToolbar",val)!=ERROR_SUCCESS || !val)
					{
						CComVariant name(L"{553891B7-A0D5-4526-BE18-D3CE461D6310}");
						CComVariant show(true);
						if (SUCCEEDED(m_pWebBrowser->ShowBrowserBar(&name,&show,NULL)))
							regKey.SetDWORDValue(L"ShowedToolbar",1);
					}
				}
			}

			HWND status;
			if (m_pBrowser && SUCCEEDED(m_pBrowser->GetControlWindow(FCW_STATUS,&status)))
			{
				m_TopWindow=GetAncestor(status,GA_ROOT);
				if (!GetProp(m_TopWindow,g_LoadedSettingsAtom))
				{
					SetProp(m_TopWindow,g_LoadedSettingsAtom,(HANDLE)1);
					LoadSettings();
				}
				bool bWin7=(LOWORD(GetVersion())==0x0106);
				bool bWin8=(LOWORD(GetVersion())==0x0206);

				m_UpButtonIndex=bWin8?0:GetSettingInt(L"ShowUpButton");
				bool bShowCaption=!bWin8 && GetSettingBool(L"ShowCaption");
				bool bShowIcon=!bWin8 && GetSettingBool(L"ShowIcon");

				if (m_TopWindow && (bShowCaption || bShowIcon))
				{
					// show the title and the icon for the main Explorer window
					DWORD flags=WTNCA_NODRAWCAPTION|WTNCA_NODRAWICON;
					if (bShowCaption) flags&=~WTNCA_NODRAWCAPTION;
					if (bShowIcon) flags&=~WTNCA_NODRAWICON;
					SetWindowThemeNonClientAttributes(m_TopWindow,WTNCA_NODRAWCAPTION|WTNCA_NODRAWICON,flags);
				}

				// find the TravelBand and the rebar
				HWND band=NULL, rebar=NULL;
				if (m_TopWindow)
					band=FindChildWindow(m_TopWindow,L"TravelBand");
				if (band)
					rebar=GetParent(band);

				bool bRedrawRebar=false;

				m_Rebar=rebar;
				m_AltD=0;
				if (rebar && GetSettingBool(L"HideSearch"))
				{
					// to remove the Search box, first find the band with ID=2 (or 4 for Win8). Then disable the child control and hide the band
					int idx=(int)SendMessage(rebar,RB_IDTOINDEX,bWin8?4:2,0);
					if (idx>=0)
					{
						REBARBANDINFO info={sizeof(info),RBBIM_CHILD};
						SendMessage(rebar,RB_GETBANDINFO,idx,(LPARAM)&info);
						if (info.hwndChild)
							EnableWindow(info.hwndChild,FALSE);
						SendMessage(rebar,RB_SHOWBAND,idx,FALSE);
						bRedrawRebar=true;
					}
				}

				if (rebar && m_UpButtonIndex)
				{
					// find the toolbar 
					HWND toolbar=FindWindowEx(band,NULL,TOOLBARCLASSNAME,NULL);
					RECT rc;
					GetClientRect(toolbar,&rc);
					bool bDef;
					int size=GetSettingInt(L"UpIconSize",bDef);
					if (bDef)
						size=rc.bottom;
					m_Toolbar.Create(rebar,NULL,NULL,WS_CHILD|TBSTYLE_TOOLTIPS|TBSTYLE_FLAT|TBSTYLE_CUSTOMERASE|CCS_NODIVIDER|CCS_NOPARENTALIGN|CCS_NORESIZE);
					m_Toolbar.SendMessage(TB_SETEXTENDEDSTYLE,0,TBSTYLE_EX_MIXEDBUTTONS);
					m_Toolbar.SendMessage(TB_BUTTONSTRUCTSIZE,sizeof(TBBUTTON));
					m_Toolbar.SendMessage(TB_SETMAXTEXTROWS,1);

					std::vector<HMODULE> modules;
					CString str=GetSettingString(L"UpIconNormal");
					m_IconNormal=str?LoadIcon(size,str,modules):NULL;
					if (m_IconNormal)
					{
						str=GetSettingString(L"UpIconHot");
						m_IconHot=LoadIcon(size,str,modules);
						str=GetSettingString(L"UpIconPressed");
						m_IconPressed=LoadIcon(size,str,modules);
						str=GetSettingString(L"UpIconDisabled");
						m_IconDisabled=LoadIcon(size,str,modules);
						if (!m_IconDisabled)
							m_IconDisabled=CreateDisabledIcon(m_IconNormal,size);
					}
					else
					{
						m_IconNormal=(HICON)LoadImage(g_Instance,MAKEINTRESOURCE(IDI_UP2NORMAL),IMAGE_ICON,size,size,LR_DEFAULTCOLOR);
						m_IconHot=(HICON)LoadImage(g_Instance,MAKEINTRESOURCE(IDI_UP2HOT),IMAGE_ICON,size,size,LR_DEFAULTCOLOR);
						m_IconPressed=(HICON)LoadImage(g_Instance,MAKEINTRESOURCE(IDI_UP2PRESSED),IMAGE_ICON,size,size,LR_DEFAULTCOLOR);
						m_IconDisabled=(HICON)LoadImage(g_Instance,MAKEINTRESOURCE(IDI_UP2DISABLED),IMAGE_ICON,size,size,LR_DEFAULTCOLOR);
					}

					for (std::vector<HMODULE>::const_iterator it=modules.begin();it!=modules.end();++it)
						FreeLibrary(*it);

					TBBUTTON button={I_IMAGENONE,1,TBSTATE_ENABLED};
					m_Toolbar.SendMessage(TB_ADDBUTTONS,1,(LPARAM)&button);
					m_Toolbar.SendMessage(TB_SETBUTTONSIZE,0,MAKELONG(size,size));

					m_bRemapBands=true;
					SetWindowSubclass(rebar,RebarSubclassProc,(UINT_PTR)this,(DWORD_PTR)m_Toolbar.m_hWnd);
					REBARBANDINFO info={sizeof(info),RBBIM_CHILD|RBBIM_ID|RBBIM_CHILDSIZE|RBBIM_IDEALSIZE|RBBIM_SIZE|RBBIM_STYLE};
					info.fStyle=RBBS_HIDETITLE|RBBS_NOGRIPPER|RBBS_FIXEDSIZE;
					info.hwndChild=m_Toolbar.m_hWnd;
					info.cxIdeal=info.cx=info.cxMinChild=size;
					info.cyMinChild=size;
					info.wID='UBTN';
					SendMessage(rebar,RB_INSERTBAND,m_UpButtonIndex-1,(LPARAM)&info);
					bRedrawRebar=true;
				}

				if (rebar)
				{
					int AddressBarHistory=GetSettingInt(L"AddressBarHistory");
					HWND progress=NULL;
					HWND breadcrumbs=FindChildWindow(rebar,L"Breadcrumb Parent");
					if (breadcrumbs)
					{
						progress=GetParent(breadcrumbs);
						breadcrumbs=FindWindowEx(breadcrumbs,NULL,TOOLBARCLASSNAME,NULL);
						m_AltD=(char)GetSettingString(L"AddressAltD")[0];
						if (m_AltD>='a' && m_AltD<='z')
							m_AltD+='A'-'a';
						if (m_AltD<'A' || m_AltD>'Z')
							m_AltD=0;
					}
					m_Breadcrumbs=breadcrumbs;
					if (GetSettingBool(L"DisableBreadcrumbs"))
					{
						// "hide" the breadcrumbs. no, not really. instead of hiding the breadcrumbs we just make them show the full path as text
						if (breadcrumbs)
						{
							m_bNoBreadcrumbs=true;
							SetWindowSubclass(breadcrumbs,BreadcrumbSubclassProc,(UINT_PTR)this,0);
						}
					}
					if (progress && AddressBarHistory)
					{
						m_ComboBox=FindWindowEx(progress,NULL,WC_COMBOBOXEX,NULL);
						SetWindowSubclass(progress,ProgressSubclassProc,(UINT_PTR)this,AddressBarHistory);
						m_NavigateMsg=RegisterWindowMessage(L"ClassicShell.Navigate");
					}
				}

				if (bRedrawRebar)
					RedrawWindow(rebar,NULL,NULL,RDW_UPDATENOW|RDW_ALLCHILDREN);

				m_UpHotkey=GetSettingInt(L"UpHotkey");
				if ((m_AltD || m_UpHotkey) && !m_HookKbd)
				{
					m_HookKbd=SetWindowsHookEx(WH_KEYBOARD,HookKeyboard,NULL,GetCurrentThreadId());
				}

				if (!bWin8 && GetSettingBool(L"ShowFreeSpace"))
				{
					DWORD FreeSpace=GetSettingBool(L"ShowInfoTip")?SPACE_INFOTIP:0;
					if (bWin7)
						FreeSpace|=SPACE_WIN7;
					SetWindowSubclass(status,SubclassStatusProc,(UINT_PTR)this,FreeSpace);
					m_bForceRefresh=(bWin7 && GetSettingBool(L"ForceRefreshWin7"));
				}
			}
			s_AutoNavDelay=GetSettingInt(L"AutoNavDelay");
		}
	}
	else
	{
		// unhook
		GetTlsData()->bho=NULL;
		if (m_Hook)
			UnhookWindowsHookEx(m_Hook);
		m_Hook=NULL;
		if (m_HookKbd)
			UnhookWindowsHookEx(m_HookKbd);
		m_HookKbd=NULL;
		m_Breadcrumbs=NULL;
		if (m_Rebar)
			RemoveWindowSubclass(m_Rebar,RebarSubclassProc,(UINT_PTR)this);
		m_pBrowser=NULL;
		if (m_pWebBrowser && m_dwEventCookie!=0xFEFEFEFE)
			DispEventUnadvise(m_pWebBrowser,&DIID_DWebBrowserEvents2);
		m_pWebBrowser=NULL;
		if (m_Toolbar.m_hWnd)
			m_Toolbar.DestroyWindow();
		if (m_IconNormal) DestroyIcon(m_IconNormal); m_IconNormal=NULL;
		if (m_IconHot) DestroyIcon(m_IconHot); m_IconHot=NULL;
		if (m_IconPressed) DestroyIcon(m_IconPressed); m_IconPressed=NULL;
		if (m_IconDisabled) DestroyIcon(m_IconDisabled); m_IconDisabled=NULL;
		if (m_CurIcon) DestroyIcon(m_CurIcon); m_CurIcon=NULL;
		if (m_CurPidl) ILFree(m_CurPidl); m_CurPidl=NULL;
		if (m_NavigatePidl) ILFree(m_NavigatePidl); m_NavigatePidl=NULL;
		if (m_TopWindow) RemoveProp(m_TopWindow,g_LoadedSettingsAtom);
		m_TopWindow=NULL;
	}
	ClearComboItems();
	return S_OK;
}

void CExplorerBHO::ClearComboItems( void )
{
	for (std::vector<ComboItem>::iterator it=m_ComboItems.begin();it!=m_ComboItems.end();++it)
		ILFree(it->pidl);
	m_ComboItems.clear();
	if (m_ComboBox.m_hWnd)
	{
		for (int i=(int)m_ComboBox.SendMessage(CB_GETCOUNT)-1;i>=0;i--)
			m_ComboBox.SendMessage(CBEM_DELETEITEM,i);
	}
}

STDMETHODIMP CExplorerBHO::OnNavigateComplete( IDispatch *pDisp, VARIANT *URL )
{
	// this is called when the current folder changes. disable the Up button if this is the desktop folder
	bool bDesktop=false;
	m_CurPath[0]=0;
	if (m_CurIcon) DestroyIcon(m_CurIcon);
	m_CurIcon=NULL;
	if (m_CurPidl) ILFree(m_CurPidl);
	m_CurPidl=NULL;
	if (m_pBrowser)
	{
		CComPtr<IShellView> pView;
		m_pBrowser->QueryActiveShellView(&pView);
		if (pView)
		{
			CComQIPtr<IFolderView> pFolderView=pView;
			if (pFolderView)
			{
				if (GetSettingBool(L"ShowHeaders"))
				{

					// ***********************************************************************

					// The code to turn on the headers is borrowed from the Explorer7Fixes project under the terms of the MIT license:
					// http://github.com/ijprest/Explorer7Fixes - Copyright (c) 2010 Ian Prest

					CComQIPtr<IFolderView2> pView2=pFolderView;
					// Turn on the sort header!
					pView2->SetCurrentFolderFlags(FWF_NOHEADERINALLVIEWS,0);
					// It seems the ItemsView doesn't respect the FWF_NOHEADERINALLVIEWS flag
					// until the view has been refreshed.  Rather than call Refresh, we just
					// briefly change the view mode and change it back.
					FOLDERVIEWMODE viewMode;
					int itemSize=0;
					pView2->GetViewModeAndIconSize(&viewMode,&itemSize);
					if (viewMode!=FVM_DETAILS)
					{
						pView2->SetViewModeAndIconSize(viewMode==FVM_LIST?FVM_SMALLICON:FVM_LIST,itemSize);
						pView2->SetViewModeAndIconSize(viewMode,itemSize);
					}
					// ***********************************************************************

				}

				CComPtr<IPersistFolder2> pFolder;
				pFolderView->GetFolder(IID_IPersistFolder2,(void**)&pFolder);
				if (pFolder && SUCCEEDED(pFolder->GetCurFolder(&m_CurPidl)) && m_CurPidl)
				{
					if (ILIsEmpty(m_CurPidl))
						bDesktop=true; // only the top level has empty PIDL

					// find path and icon
					// it is possible to get the path and icon from the caption of the main window, but there are 2 problems:
					//   1) on Vista the icon is wrong. after you navigate to a new folder the icon switches to some default image and doesn't change any more
					//   2) if the user has not checked "display full path in title bar", the caption of the main window is just the current folder name and not the full path
					// so do it the hard way and grab it from SHGetNameFromIDList and SHGetFileInfo
					if (m_bNoBreadcrumbs)
					{
						wchar_t *pPath=NULL;
						if (SUCCEEDED(SHGetNameFromIDList(m_CurPidl,SIGDN_DESKTOPABSOLUTEEDITING,&pPath)))
							Strcpy(m_CurPath,_countof(m_CurPath),pPath);
						else if (SUCCEEDED(SHGetNameFromIDList(m_CurPidl,SIGDN_FILESYSPATH,&pPath)))
							Strcpy(m_CurPath,_countof(m_CurPath),pPath); // just in case DESKTOPABSOLUTE fails let's try the FILESYSPATH. probably not needed
						if (pPath) CoTaskMemFree(pPath);
						SHFILEINFO info;
						if (SUCCEEDED(SHGetFileInfo((LPCTSTR)m_CurPidl,0,&info,sizeof(info),SHGFI_ICON|SHGFI_SMALLICON|SHGFI_PIDL)))
							m_CurIcon=info.hIcon;
					}
				}
				else
				{
					m_CurPidl=NULL;
				}
			}
		}
		if (GetSettingBool(L"HideScrollTip"))
		{
			CComQIPtr<IFolderViewOptions> pOptions=m_pBrowser;
			if (pOptions)
				pOptions->SetFolderViewOptions(FVO_NOSCROLLTIPS,FVO_NOSCROLLTIPS);
		}
	}
	if (m_Toolbar.m_hWnd)
		m_Toolbar.SendMessage(TB_ENABLEBUTTON,1,bDesktop?0:1);
	return S_OK;
}

STDMETHODIMP CExplorerBHO::OnQuit( void )
{
	if (m_pWebBrowser && m_dwEventCookie!=0xFEFEFEFE) // ATL's event cookie is 0xFEFEFEFE, when the sink is not advised
		return DispEventUnadvise(m_pWebBrowser,&DIID_DWebBrowserEvents2);
	return S_OK;
}

bool ShowTreeProperties( HWND hwndTree )
{
	PIDLIST_ABSOLUTE pidl=GetSelectedItem(hwndTree);
	if (pidl)
	{
		// show properties
		SHELLEXECUTEINFO execute={sizeof(execute),SEE_MASK_IDLIST|SEE_MASK_INVOKEIDLIST,NULL,L"properties"};
		execute.lpIDList=pidl;
		execute.nShow=SW_SHOWNORMAL;
		ShellExecuteEx(&execute);
		ILFree(pidl);
		return true;
	}
	return false;
}
