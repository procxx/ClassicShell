// Classic Shell (c) 2009-2010, Ivo Beltchev
// The sources for Classic Shell are distributed under the MIT open source license

// ExplorerBHO.cpp : Implementation of CExplorerBHO

#include "stdafx.h"
#include "ExplorerBHO.h"
#include "Settings.h"
#include "GlobalSettings.h"
#include "TranslationSettings.h"
#include "dllmain.h"
#include <uxtheme.h>
#include <dwmapi.h>
#include <Ntquery.h>

// CExplorerBHO - a browser helper object that implements Alt+Enter for the folder tree

__declspec(thread) HHOOK CExplorerBHO::s_Hook; // one hook per thread

struct FindChild
{
	const wchar_t *className;
	HWND hWnd;
};

static BOOL CALLBACK EnumChildProc( HWND hwnd, LPARAM lParam )
{
	FindChild &find=*(FindChild*)lParam;
	wchar_t name[256];
	GetClassName(hwnd,name,_countof(name));
	if (_wcsicmp(name,find.className)!=0) return TRUE;
	find.hWnd=hwnd;
	return FALSE;
}

static HWND FindChildWindow( HWND hwnd, const wchar_t *className )
{
	FindChild find={className};
	EnumChildWindows(hwnd,EnumChildProc,(LPARAM)&find);
	return find.hWnd;
}

const UINT_PTR TIMER_NAVIGATE='CLSH';

static LRESULT CALLBACK SubclassParentProc( HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData )
{
	// when the tree selection changes start a timer to navigate to the new folder in 100ms
	if (uMsg==WM_NOTIFY && ((NMHDR*)lParam)->code==TVN_SELCHANGED && ((NMTREEVIEW*)lParam)->action==TVC_BYKEYBOARD)
		SetTimer(((NMHDR*)lParam)->hwndFrom,TIMER_NAVIGATE,100,NULL);
	return DefSubclassProc(hWnd,uMsg,wParam,lParam);
}

static LRESULT CALLBACK SubclassTreeProc( HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData )
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
		// time to navigate to the new folder (simulate pressing Space)
		PostMessage(hWnd,WM_KEYDOWN,VK_SPACE,0);
		KillTimer(hWnd,TIMER_NAVIGATE);
		return 0;
	}

	if (uMsg==WM_CHAR && wParam==' ')
		return 0; // ignore the Space character (to stop the tree view from beeping)

	if (uMsg==WM_SYSKEYDOWN && wParam==VK_RETURN)
	{
		// Alt+Enter is pressed
		// if this message was for the folder tree, show the properties of the selected item
		DWORD FoldersSettings=CExplorerBHO::FOLDERS_DEFAULT;
		CRegKey regSettings;
		if (regSettings.Open(HKEY_CURRENT_USER,L"Software\\IvoSoft\\ClassicExplorer")==ERROR_SUCCESS)
			regSettings.QueryDWORDValue(L"FoldersSettings",FoldersSettings);

		if ((FoldersSettings&CExplorerBHO::FOLDERS_ALTENTER) && ShowTreeProperties(hWnd))
			return 0;
	}
	if (uMsg==TVM_SETEXTENDEDSTYLE && wParam==(TVS_EX_FADEINOUTEXPANDOS|TVS_EX_AUTOHSCROLL|0x80000000) && lParam==0)
	{
		wParam&=0x7FFFFFFF;

		DWORD FoldersSettings=CExplorerBHO::FOLDERS_DEFAULT;
		CRegKey regSettings;
		if (regSettings.Open(HKEY_CURRENT_USER,L"Software\\IvoSoft\\ClassicExplorer")==ERROR_SUCCESS)
			regSettings.QueryDWORDValue(L"FoldersSettings",FoldersSettings);

		if (FoldersSettings&CExplorerBHO::FOLDERS_AUTONAVIGATE)
			SetWindowSubclass(GetParent(hWnd),SubclassParentProc,'CLSH',0);

		if (!(FoldersSettings&CExplorerBHO::FOLDERS_NOFADE))
			wParam&=~TVS_EX_FADEINOUTEXPANDOS;

		int indent=-1;
		if (FoldersSettings&CExplorerBHO::FOLDERS_FULLINDENT)
			indent=0;

		if ((FoldersSettings&CExplorerBHO::FOLDERS_STYLE_MASK)!=CExplorerBHO::FOLDERS_VISTA)
		{
			SetWindowTheme(hWnd,NULL,NULL);
			DWORD style=GetWindowLong(hWnd,GWL_STYLE);
			style&=~TVS_NOHSCROLL;
			if ((FoldersSettings&CExplorerBHO::FOLDERS_STYLE_MASK)==CExplorerBHO::FOLDERS_SIMPLE)
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
			SetWindowLong(hWnd,GWL_STYLE,style);
		}
		else
		{
			wParam&=~TVS_EX_AUTOHSCROLL;
		}
		if (indent>=0)
			TreeView_SetIndent(hWnd,indent);

		if (wParam==0) return 0;
	}
	return DefSubclassProc(hWnd,uMsg,wParam,lParam);
}

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

			// find the current folder and show the free space of the drive containing the current folder
			// also works for network locations
			IShellBrowser *pBrowser=((CExplorerBHO*)uIdSubclass)->m_pBrowser;
			CComPtr<IShellView> pView;
			if (pBrowser && SUCCEEDED(pBrowser->QueryActiveShellView(&pView)))
			{
				CComQIPtr<IFolderView> pView2=pView;
				CComPtr<IPersistFolder2> pFolder;
				if (pView2 && SUCCEEDED(pView2->GetFolder(IID_IPersistFolder2,(void**)&pFolder)))
				{
					LPITEMIDLIST pidl;
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
								swprintf_s(buf,FindTranslation("Status.FreeSpace",L"%s (Disk free space: %s)"),text,str);
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
			LPITEMIDLIST pidl;
			if (pView2 && SUCCEEDED(pView2->GetFolder(IID_IPersistFolder2,(void**)&pFolder)) && SUCCEEDED(pFolder->GetCurFolder(&pidl)))
			{
				CComQIPtr<IShellFolder2> pFolder2=pFolder;
				UINT type=SVGIO_SELECTION;
				int count;
				if ((dwRefData&SPACE_TOTAL) && (FAILED(pView2->ItemCount(SVGIO_SELECTION,&count)) || count==0))
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

LRESULT CALLBACK CExplorerBHO::HookExplorer( int nCode, WPARAM wParam, LPARAM lParam )
{
	if (nCode==HCBT_CREATEWND)
	{
		HWND hWnd=(HWND)wParam;
		CBT_CREATEWND *create=(CBT_CREATEWND*)lParam;
		if (create->lpcs->lpszClass>(LPTSTR)0xFFFF && _wcsicmp(create->lpcs->lpszClass,WC_TREEVIEW)==0)
		{
			DWORD_PTR settings=0;
			if (LOWORD(GetVersion())==0x0106 && FindSettingBool("FixFolderScroll",true))
				settings|=1;
			SetWindowSubclass(hWnd,SubclassTreeProc,'CLSH',settings);
			PostMessage(hWnd,TVM_SETEXTENDEDSTYLE,TVS_EX_FADEINOUTEXPANDOS|TVS_EX_AUTOHSCROLL|0x80000000,0);
			UnhookWindowsHookEx(s_Hook);
			s_Hook=NULL;
			return 0;
		}
	}
	return CallNextHookEx(NULL,nCode,wParam,lParam);
}

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
		wcscpy_s(pTip->pszText,pTip->cchTextMax,FindTranslation("Toolbar.GoUp",L"Up One Level"));
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
	return DefSubclassProc(hWnd,uMsg,wParam,lParam);
}

HRESULT STDMETHODCALLTYPE CExplorerBHO::SetSite( IUnknown *pUnkSite )
{
	IObjectWithSiteImpl<CExplorerBHO>::SetSite(pUnkSite);

	if (pUnkSite)
	{
		ReadIniFile(false);

		// hook
		if (!s_Hook)
		{
			s_Hook=SetWindowsHookEx(WH_CBT,HookExplorer,NULL,GetCurrentThreadId());
		}
		CComQIPtr<IServiceProvider> pProvider=pUnkSite;

		if (pProvider)
		{
			pProvider->QueryService(SID_SShellBrowser,IID_IShellBrowser,(void**)&m_pBrowser);

			HWND status;
			if (m_pBrowser && SUCCEEDED(m_pBrowser->GetControlWindow(FCW_STATUS,&status)))
			{
				CRegKey regSettings;
				bool bWin7=(LOWORD(GetVersion())==0x0106);
				DWORD FreeSpace=bWin7?SPACE_SHOW:0;
				DWORD UpButton=1;
				if (regSettings.Open(HKEY_CURRENT_USER,L"Software\\IvoSoft\\ClassicExplorer")==ERROR_SUCCESS)
				{
					regSettings.QueryDWORDValue(L"FreeSpace",FreeSpace);
					regSettings.QueryDWORDValue(L"UpButton",UpButton);
				}

				if (UpButton)
				{
					// find the TravelBand, the rebar and the toolbar
					HWND band=GetAncestor(status,GA_ROOT);
					if (band)
						band=FindChildWindow(band,L"TravelBand");
					if (band)
					{
						HWND toolbar=FindWindowEx(band,NULL,TOOLBARCLASSNAME,NULL);
						RECT rc;
						GetClientRect(toolbar,&rc);
						HWND rebar=GetParent(band);
						const wchar_t *str=FindSetting("UpIconSize");
						int size=str?_wtol(str):rc.bottom;
						m_Toolbar=CreateWindow(TOOLBARCLASSNAME,L"",WS_CHILD|TBSTYLE_TOOLTIPS|TBSTYLE_FLAT|TBSTYLE_CUSTOMERASE|CCS_NODIVIDER|CCS_NOPARENTALIGN|CCS_NORESIZE,0,0,10,10,rebar,NULL,g_Instance,NULL);
						m_Toolbar.SendMessage(TB_SETEXTENDEDSTYLE,0,TBSTYLE_EX_MIXEDBUTTONS);
						m_Toolbar.SendMessage(TB_BUTTONSTRUCTSIZE,sizeof(TBBUTTON));
						m_Toolbar.SendMessage(TB_SETMAXTEXTROWS,1);

						HMODULE hShell32=GetModuleHandle(L"Shell32.dll");
						std::vector<HMODULE> modules;
						str=FindSetting("UpIconNormal");
						m_IconNormal=str?LoadIcon(size,str,0,modules,hShell32):NULL;
						if (m_IconNormal)
						{
							str=FindSetting("UpIconHot");
							m_IconHot=str?LoadIcon(size,str,0,modules,NULL):NULL;
							str=FindSetting("UpIconPressed");
							m_IconPressed=str?LoadIcon(size,str,0,modules,NULL):NULL;
							str=FindSetting("UpIconDisabled");
							m_IconDisabled=str?LoadIcon(size,str,0,modules,NULL):NULL;
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

						SetWindowSubclass(rebar,RebarSubclassProc,(UINT_PTR)this,(DWORD_PTR)m_Toolbar.m_hWnd);
						REBARBANDINFO info={sizeof(info),RBBIM_CHILD|RBBIM_CHILDSIZE|RBBIM_IDEALSIZE|RBBIM_SIZE|RBBIM_STYLE};
						SendMessage(rebar,RB_GETBANDINFO,1,(LPARAM)&info);
						info.fStyle=RBBS_HIDETITLE|RBBS_NOGRIPPER|RBBS_FIXEDSIZE;
						info.hwndChild=m_Toolbar.m_hWnd;
						info.cxIdeal=info.cx=info.cxMinChild=size;
						info.cyMinChild=size;
						SendMessage(rebar,RB_INSERTBAND,1,(LPARAM)&info);
						RedrawWindow(rebar,NULL,NULL,RDW_UPDATENOW|RDW_ALLCHILDREN);

						// listen for web browser notifications. we only care about DISPID_DOWNLOADCOMPLETE and DISPID_ONQUIT
						pProvider->QueryService(SID_SWebBrowserApp,IID_IWebBrowser2,(void**)&m_pWebBrowser);
						if (m_pWebBrowser)
						{
							if (m_dwEventCookie==0xFEFEFEFE) // ATL's event cookie is 0xFEFEFEFE when the sink is not advised
								DispEventAdvise(m_pWebBrowser,&DIID_DWebBrowserEvents2);
						}
					}
				}

				if (FreeSpace)
				{
					FreeSpace=SPACE_SHOW|SPACE_TOTAL; // always show total
					if (bWin7)
						FreeSpace|=SPACE_WIN7;
					SetWindowSubclass(status,SubclassStatusProc,(UINT_PTR)this,FreeSpace);
					m_bForceRefresh=(bWin7 && FindSettingBool("ForceRefreshWin7",true));
				}
			}
		}
	}
	else
	{
		// unhook
		if (s_Hook)
			UnhookWindowsHookEx(s_Hook);
		s_Hook=NULL;
		m_pBrowser=NULL;
		if (m_pWebBrowser && m_dwEventCookie!=0xFEFEFEFE)
			DispEventUnadvise(m_pWebBrowser,&DIID_DWebBrowserEvents2);
		m_pWebBrowser=NULL;
		m_Toolbar=NULL;
		if (m_IconNormal) DestroyIcon(m_IconNormal); m_IconNormal=NULL;
		if (m_IconHot) DestroyIcon(m_IconHot); m_IconHot=NULL;
		if (m_IconPressed) DestroyIcon(m_IconPressed); m_IconPressed=NULL;
		if (m_IconDisabled) DestroyIcon(m_IconDisabled); m_IconDisabled=NULL;
	}
	return S_OK;
}


STDMETHODIMP CExplorerBHO::OnNavigateComplete( IDispatch *pDisp, VARIANT *URL )
{
	// this is called when the current folder changes. disable the Up button if this is the desktop folder
	bool bDesktop=false;
	if (m_pBrowser)
	{
		CComPtr<IShellView> pView;
		m_pBrowser->QueryActiveShellView(&pView);
		if (pView)
		{
			CComQIPtr<IFolderView> pView2=pView;
			if (pView2)
			{
				CComPtr<IPersistFolder2> pFolder;
				pView2->GetFolder(IID_IPersistFolder2,(void**)&pFolder);
				if (pFolder)
				{
					LPITEMIDLIST pidl;
					pFolder->GetCurFolder(&pidl);
					if (ILIsEmpty(pidl))
						bDesktop=true; // only the top level has empty PIDL
					CoTaskMemFree(pidl);
				}
			}
		}
	}
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
	// find the PIDL of the selected item (combine all child PIDLs from the current item and its parents)
	HTREEITEM hItem=TreeView_GetSelection(hwndTree);
	LPITEMIDLIST pidl=NULL;
	while (hItem)
	{
		TVITEMEX info={TVIF_PARAM,hItem};
		TreeView_GetItem(hwndTree,&info);
		LPITEMIDLIST **pidl1=(LPITEMIDLIST**)info.lParam;
		if (!pidl1 || !*pidl1 || !**pidl1)
		{
			if (pidl) ILFree(pidl);
			pidl=NULL;
			break;
		}
		LPITEMIDLIST pidl2=pidl?ILCombine(**pidl1,pidl):ILClone(**pidl1);
		if (pidl) ILFree(pidl);
		pidl=pidl2;
		hItem=TreeView_GetParent(hwndTree,hItem);
	}
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
