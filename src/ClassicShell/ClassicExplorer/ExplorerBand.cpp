// Classic Shell (c) 2009, Ivo Beltchev
// The sources for Classic Shell are distributed under the MIT open source license

// ExplorerBand.cpp : Implementation of CExplorerBand

#include "stdafx.h"
#include "ExplorerBand.h"
#include "resource.h"
#include "..\LocalizationSettings\ParseSettings.h"

// Dialog proc for the Settings dialog. Edits the settings and saves them to the registry
INT_PTR CALLBACK SettingsDlgProc( HWND hwndDlg, UINT uMsg, WPARAM wParam, LPARAM lParam )
{
	if (uMsg==WM_INITDIALOG)
	{
		CRegKey regSettings;
		if (regSettings.Open(HKEY_CURRENT_USER,L"Software\\IvoSoft\\ClassicExplorer")!=ERROR_SUCCESS)
			regSettings.Create(HKEY_CURRENT_USER,L"Software\\IvoSoft\\ClassicExplorer");

		DWORD EnableCopyUI, EnableAltEnter;
		if (regSettings.QueryDWORDValue(L"EnableCopyUI",EnableCopyUI)!=ERROR_SUCCESS)
			EnableCopyUI=1;
		if (regSettings.QueryDWORDValue(L"EnableAltEnter",EnableAltEnter)!=ERROR_SUCCESS)
			EnableAltEnter=1;

		RECT rc1,rc2;
		GetWindowRect(hwndDlg,&rc1);
		GetWindowRect(GetParent(hwndDlg),&rc2);
		OffsetRect(&rc1,(rc2.left+rc2.right)/2-(rc1.left+rc1.right)/2,(rc2.top+rc2.bottom)/2-(rc1.top+rc1.bottom)/2);
		if (rc1.top<rc2.top) OffsetRect(&rc1,0,rc2.top-rc1.top);
		SetWindowPos(hwndDlg,NULL,rc1.left,rc1.top,rc1.right-rc1.left,rc1.bottom-rc1.top,SWP_NOZORDER);
		SendMessage(hwndDlg,DM_REPOSITION,0,0);

		CheckDlgButton(hwndDlg,IDC_CHECKCOPY,EnableCopyUI?BST_CHECKED:BST_UNCHECKED);
		CheckDlgButton(hwndDlg,IDC_CHECKBHO,EnableAltEnter?BST_CHECKED:BST_UNCHECKED);
		return TRUE;
	}
	if (uMsg==WM_COMMAND && wParam==IDOK)
	{
		CRegKey regSettings;
		if (regSettings.Open(HKEY_CURRENT_USER,L"Software\\IvoSoft\\ClassicExplorer")!=ERROR_SUCCESS)
			regSettings.Create(HKEY_CURRENT_USER,L"Software\\IvoSoft\\ClassicExplorer");

		DWORD EnableCopyUI, EnableAltEnter;
		if (regSettings.QueryDWORDValue(L"EnableCopyUI",EnableCopyUI)!=ERROR_SUCCESS)
			EnableCopyUI=1;
		if (regSettings.QueryDWORDValue(L"EnableAltEnter",EnableAltEnter)!=ERROR_SUCCESS)
			EnableAltEnter=1;

		DWORD EnableCopyUI2=(IsDlgButtonChecked(hwndDlg,IDC_CHECKCOPY)==BST_CHECKED)?1:0;
		DWORD EnableAltEnter2=(IsDlgButtonChecked(hwndDlg,IDC_CHECKBHO)==BST_CHECKED)?1:0;

		int res=0;
		if (EnableCopyUI!=EnableCopyUI2)
		{
			regSettings.SetDWORDValue(L"EnableCopyUI",EnableCopyUI2);
			res=1;
		}
		if (EnableAltEnter!=EnableAltEnter2)
		{
			regSettings.SetDWORDValue(L"EnableAltEnter",EnableAltEnter2);
			res=1;
		}

		EndDialog(hwndDlg,res);
		return TRUE;
	}
	if (uMsg==WM_COMMAND && wParam==IDCANCEL)
	{
		EndDialog(hwndDlg,0);
		return TRUE;
	}
	if (uMsg==WM_NOTIFY)
	{
		NMHDR *pHdr=(NMHDR*)lParam;
		if (pHdr->idFrom==IDC_LINKHELP && (pHdr->code==NM_CLICK || pHdr->code==NM_RETURN))
		{
			wchar_t path[_MAX_PATH];
			GetModuleFileName(g_Instance,path,_countof(path));
			*PathFindFileName(path)=0;
			wcscat_s(path,L"ClassicExplorer.html");
			ShellExecute(NULL,NULL,path,NULL,NULL,SW_SHOWNORMAL);
			return TRUE;
		}
	}
	return FALSE;
}

///////////////////////////////////////////////////////////////////////////////

// CBandWindow - the parent window of the toolbar

LRESULT CBandWindow::OnCreate( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled )
{
	// create the toolbar
	m_hWndToolbar=CreateWindow(TOOLBARCLASSNAME,L"",WS_CHILD|TBSTYLE_TOOLTIPS|TBSTYLE_FLAT|TBSTYLE_LIST|CCS_NODIVIDER|CCS_NOPARENTALIGN|CCS_NORESIZE,0,0,10,10,m_hWnd,(HMENU)101,g_Instance,NULL);

	::SendMessage(m_hWndToolbar,TB_SETEXTENDEDSTYLE,0,TBSTYLE_EX_MIXEDBUTTONS);
	::SendMessage(m_hWndToolbar,TB_BUTTONSTRUCTSIZE,sizeof(TBBUTTON),0);
	::SendMessage(m_hWndToolbar, TB_SETMAXTEXTROWS, 1, 0L);

	// pick icon size based on the DPI setting
	HDC hdc=::GetDC(NULL);
	int dpi=GetDeviceCaps(hdc,LOGPIXELSY);
	::ReleaseDC(NULL,hdc);
	int iconSize;
	if (dpi>120)
		iconSize=24;
	else if (dpi>96)
		iconSize=20;
	else
		iconSize=16;

	m_Enabled=ImageList_Create(iconSize,iconSize,ILC_COLOR32|(IsLanguageRTL()?ILC_MIRROR:0),0,2);

	// load icons for parent folder (46) and folder settings (210) from Shell32.dll
	HMODULE hModule=GetModuleHandle(L"Shell32.dll");

	HICON iconUp=(HICON)LoadImage(hModule,MAKEINTRESOURCE(46),IMAGE_ICON,iconSize,iconSize,LR_DEFAULTCOLOR);
	ImageList_AddIcon(m_Enabled,iconUp);
	DestroyIcon(iconUp);

	HICON iconSettings=(HICON)LoadImage(hModule,MAKEINTRESOURCE(210),IMAGE_ICON,iconSize,iconSize,LR_DEFAULTCOLOR);
	ImageList_AddIcon(m_Enabled,iconSettings);
	DestroyIcon(iconSettings);

	HIMAGELIST old=(HIMAGELIST)::SendMessage(m_hWndToolbar,TB_SETIMAGELIST,0,(LPARAM)m_Enabled);
	if (old) ImageList_Destroy(old);

	// add buttons
	TBBUTTON buttons[2]={
		{0,ID_GOUP,TBSTATE_ENABLED,BTNS_BUTTON|BTNS_AUTOSIZE,{0},0,(INT_PTR)FindSetting("Toolbar.GoUp",L"Up One Level")},
		{1,ID_SETTINGS,TBSTATE_ENABLED,BTNS_BUTTON|BTNS_AUTOSIZE,{0},0,(INT_PTR)FindSetting("Toolbar.Settings",L"Classic Explorer Settings")},
	};
	::SendMessage(m_hWndToolbar,TB_ADDBUTTONS,_countof(buttons),(LPARAM)buttons);
	return 0;
}

LRESULT CBandWindow::OnDestroy( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled )
{
	ImageList_Destroy(m_Enabled);
	bHandled=FALSE;
	return 0;
}

// Go to the parent folder
LRESULT CBandWindow::OnGoUp( WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled )
{
	if (m_pBrowser)
		m_pBrowser->BrowseObject(NULL,(GetKeyState(VK_CONTROL)<0?SBSP_NEWBROWSER:SBSP_SAMEBROWSER)|SBSP_DEFMODE|SBSP_PARENT);
	return TRUE;
}

// Show the settings dialog
LRESULT CBandWindow::OnSettings( WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled )
{
#ifndef BUILD_SETUP
	if (GetKeyState(VK_SHIFT)<0)
		*(int*)0=0; // force a crash if Shift is pressed. Makes it easy to restart explorer.exe
#endif
	DialogBox(g_Instance,MAKEINTRESOURCE(IDD_SETTINGS),m_hWnd,SettingsDlgProc);
	return TRUE;
}

void CBandWindow::UpdateToolbar( void )
{
	// disable the Up button if we are at the top level
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
	SendMessage(m_hWndToolbar,TB_ENABLEBUTTON,CBandWindow::ID_GOUP,bDesktop?0:1);
}

///////////////////////////////////////////////////////////////////////////////

// CExplorerBand - adds a toolbar band to Windows Explorer with 2 buttons - "Up" and "Settings"

CExplorerBand::CExplorerBand( void )
{
	m_bSubclassRebar=(LOWORD(GetVersion())==0x0106); // Windows 7
	m_bSubclassedRebar=false;
}

// Subclasses the rebar control on Windows 7. Makes sure the RBBS_BREAK style is properly set. Windows 7 has a bug
// that forces RBBS_BREAK for every rebar band
LRESULT CALLBACK CExplorerBand::RebarSubclassProc( HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData )
{
	if (uMsg==RB_SETBANDINFO || uMsg==RB_INSERTBAND)
	{
		REBARBANDINFO *pInfo=(REBARBANDINFO*)lParam;
		if ((pInfo->hwndChild==(HWND)dwRefData) && (pInfo->fMask&RBBIM_STYLE))
		{
			if (((CExplorerBand*)uIdSubclass)->m_bBandNewLine)
				pInfo->fStyle|=RBBS_BREAK;
			else
				pInfo->fStyle&=~RBBS_BREAK;
		}
	}
	return DefSubclassProc(hWnd,uMsg,wParam,lParam);
}

// IDeskBand
STDMETHODIMP CExplorerBand::GetBandInfo( DWORD dwBandID, DWORD dwViewMode, DESKBANDINFO* pdbi )
{
	// initializes the band
	if (m_bSubclassRebar && !m_bSubclassedRebar)
	{
		CRegKey regSettings;
		if (regSettings.Open(HKEY_CURRENT_USER,L"Software\\IvoSoft\\ClassicExplorer")!=ERROR_SUCCESS)
			regSettings.Create(HKEY_CURRENT_USER,L"Software\\IvoSoft\\ClassicExplorer");

		DWORD NewLine;
		if (regSettings.QueryDWORDValue(L"NewLine",NewLine)!=ERROR_SUCCESS)
			NewLine=0;
		m_bBandNewLine=NewLine!=0;

		HWND parent=GetParent(m_BandWindow.GetToolbar());
		wchar_t className[256];
		GetClassName(parent,className,_countof(className));
		if (_wcsicmp(className,REBARCLASSNAME)==0)
		{
			SetWindowSubclass(parent,RebarSubclassProc,(UINT_PTR)this,(DWORD_PTR)m_BandWindow.GetToolbar());
			m_bSubclassedRebar=true;
		}
	}
	RECT rc;
	int count=(int)SendMessage(m_BandWindow.GetToolbar(),TB_BUTTONCOUNT,0,0);
	SendMessage(m_BandWindow.GetToolbar(),TB_GETITEMRECT,count-1,(LPARAM)&rc);

	if (pdbi)
	{
		if (pdbi->dwMask&DBIM_MINSIZE)
		{
			pdbi->ptMinSize.x=rc.right;
			pdbi->ptMinSize.y=rc.bottom;
		}
		if (pdbi->dwMask&DBIM_MAXSIZE)
		{
			pdbi->ptMaxSize.x=0; // ignored
			pdbi->ptMaxSize.y=-1;	// unlimited
		}
		if (pdbi->dwMask&DBIM_INTEGRAL)
		{
			pdbi->ptIntegral.x=0; // not sizeable
			pdbi->ptIntegral.y=0; // not sizeable
		}
		if (pdbi->dwMask&DBIM_ACTUAL)
		{
			pdbi->ptActual.x=rc.right;
			pdbi->ptActual.y=rc.bottom;
		}
		if (pdbi->dwMask&DBIM_TITLE)
		{
			*pdbi->wszTitle=0; // no title
		}
		if (pdbi->dwMask&DBIM_BKCOLOR)
		{
			//Use the default background color by removing this flag.
			pdbi->dwMask&=~DBIM_BKCOLOR;
		}
	}
	return S_OK;
}

// IOleWindow
STDMETHODIMP CExplorerBand::GetWindow( HWND* phwnd )
{
	if (!phwnd)
		return E_INVALIDARG;
	*phwnd=m_BandWindow.GetToolbar();
	return S_OK;
}

STDMETHODIMP CExplorerBand::ContextSensitiveHelp( BOOL fEnterMode )
{
	return S_OK;
}

// IDockingWindow
STDMETHODIMP CExplorerBand::CloseDW( unsigned long dwReserved )
{
	ShowDW(FALSE);
	return S_OK;
}

STDMETHODIMP CExplorerBand::ResizeBorderDW( const RECT* prcBorder, IUnknown* punkToolbarSite, BOOL fReserved )
{
	// Not used by any band object.
	return E_NOTIMPL;
}

STDMETHODIMP CExplorerBand::ShowDW( BOOL fShow )
{
	if (m_bSubclassedRebar)
	{
		// on Windows 7 get the current RBBS_BREAK state and save it in the registry to be restored later
		HWND parent=GetParent(m_BandWindow.GetToolbar());
		int n=(int)SendMessage(parent,RB_GETBANDCOUNT,0,0);
		for (int i=0;i<n;i++)
		{
			REBARBANDINFO info={sizeof(info),RBBIM_STYLE|RBBIM_CHILD};
			SendMessage(parent,RB_GETBANDINFO,i,(LPARAM)&info);
			if (info.hwndChild==m_BandWindow.GetToolbar())
			{
				m_bBandNewLine=(info.fStyle&RBBS_BREAK)!=0;
				CRegKey regSettings;
				if (regSettings.Open(HKEY_CURRENT_USER,L"Software\\IvoSoft\\ClassicExplorer")!=ERROR_SUCCESS)
					regSettings.Create(HKEY_CURRENT_USER,L"Software\\IvoSoft\\ClassicExplorer");

				regSettings.SetDWORDValue(L"NewLine",m_bBandNewLine?1:0);
				break;
			}
		}
	}
	ShowWindow(m_BandWindow.GetToolbar(),fShow?SW_SHOW:SW_HIDE);
	return S_OK;
}

// IObjectWithSite
STDMETHODIMP CExplorerBand::SetSite( IUnknown* pUnkSite )
{
	if (m_BandWindow.IsWindow())
		m_BandWindow.DestroyWindow();
	m_BandWindow.SetBrowser(NULL);
	if (m_bSubclassedRebar)
		RemoveWindowSubclass(GetParent(m_BandWindow.GetToolbar()),RebarSubclassProc,(UINT_PTR)this);
	m_bSubclassedRebar=false;

	if (m_pWebBrowser && m_dwEventCookie!=0xFEFEFEFE)
		DispEventUnadvise(m_pWebBrowser,&DIID_DWebBrowserEvents2);
	m_pWebBrowser=NULL;

	IObjectWithSiteImpl<CExplorerBand>::SetSite(pUnkSite);

	//If punkSite is not NULL, a new site is being set.
	if (pUnkSite)
	{
		//Get the parent window.
		HWND hWndParent=NULL;

		CComQIPtr<IOleWindow> pOleWindow=pUnkSite;
		if (pOleWindow)
			pOleWindow->GetWindow(&hWndParent);

		if (!IsWindow(hWndParent))
			return E_FAIL;

		m_BandWindow.Create(hWndParent,NULL,NULL,WS_CHILD);
		if (!m_BandWindow.IsWindow())
			return E_FAIL;

		CComQIPtr<IServiceProvider> pProvider=pUnkSite;

		if (pProvider)
		{
			CComPtr<IShellBrowser> pBrowser;
			pProvider->QueryService(SID_SShellBrowser,IID_IShellBrowser,(void**)&pBrowser);
			m_BandWindow.SetBrowser(pBrowser);

			// listen for web browser notifications. we only care about DISPID_DOWNLOADCOMPLETE and DISPID_ONQUIT
			pProvider->QueryService(SID_SWebBrowserApp,IID_IWebBrowser2,(void**)&m_pWebBrowser);
			if (m_pWebBrowser)
			{
				if (m_dwEventCookie==0xFEFEFEFE) // ATL's event cookie is 0xFEFEFEFE when the sink is not advised
					DispEventAdvise(m_pWebBrowser,&DIID_DWebBrowserEvents2);
			}
		}
	}
	return S_OK;
}

STDMETHODIMP CExplorerBand::OnDownloadComplete( void )
{
	// this is called when the current folder changes. disable the Up button if this is the desktop folder
	m_BandWindow.UpdateToolbar();
	return S_OK;
}

STDMETHODIMP CExplorerBand::OnQuit( void )
{
	if (m_pWebBrowser && m_dwEventCookie!=0xFEFEFEFE) // ATL's event cookie is 0xFEFEFEFE, when the sink is not advised
		return DispEventUnadvise(m_pWebBrowser,&DIID_DWebBrowserEvents2);
	return S_OK;
}
