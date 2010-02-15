// Classic Shell (c) 2009-2010, Ivo Beltchev
// The sources for Classic Shell are distributed under the MIT open source license

// ExplorerBand.cpp : Implementation of CExplorerBand

#include "stdafx.h"
#include "ExplorerBand.h"
#include "resource.h"
#include "ExplorerBHO.h"
#include "ParseSettings.h"
#include "GlobalSettings.h"
#include "TranslationSettings.h"
#include "Settings.h"
#include "dllmain.h"

///////////////////////////////////////////////////////////////////////////////

// CBandWindow - the parent window of the toolbar

const CBandWindow::StdToolbarItem CBandWindow::s_StdItems[]={
	{ID_GOUP,"Toolbar.GoUp",L"Up One Level",46,NULL,NULL,L",2",L",3"},
	{ID_CUT,"Toolbar.Cut",L"Cut",16762},
	{ID_COPY,"Toolbar.Copy",L"Copy",243},
	{ID_PASTE,"Toolbar.Paste",L"Paste",16763},
	{ID_DELETE,"Toolbar.Delete",L"Delete",240},
	{ID_PROPERTIES,"Toolbar.Properties",L"Properties",253},
	{ID_EMAIL,"Toolbar.Email",L"E-mail the selected items",265},
	{ID_SETTINGS,"Toolbar.Settings",L"Classic Explorer Settings",210,NULL,NULL,L",1"},
};

static struct
{
	const wchar_t *name;
	int id;
} g_StdCommands[]={
	{L"up",CBandWindow::ID_GOUP},
	{L"cut",CBandWindow::ID_CUT},
	{L"copy",CBandWindow::ID_COPY},
	{L"paste",CBandWindow::ID_PASTE},
	{L"delete",CBandWindow::ID_DELETE},
	{L"properties",CBandWindow::ID_PROPERTIES},
	{L"email",CBandWindow::ID_EMAIL},
	{L"settings",CBandWindow::ID_SETTINGS},
	{L"moveto",CBandWindow::ID_MOVETO},
	{L"copyto",CBandWindow::ID_COPYTO},
	{L"undo",CBandWindow::ID_UNDO},
	{L"redo",CBandWindow::ID_REDO},
	{L"selectall",CBandWindow::ID_SELECTALL},
	{L"invertselection",CBandWindow::ID_INVERT},
	{L"back",CBandWindow::ID_GOBACK},
	{L"forward",CBandWindow::ID_GOFORWARD},
	{L"refresh",CBandWindow::ID_REFRESH},
};

bool CBandWindow::ParseToolbarItem( const wchar_t *name, StdToolbarItem &item )
{
	wchar_t text[256];
	swprintf_s(text,L"%s.Command",name);
	const wchar_t *str=FindSetting(text);
	if (!str) return false;

	item.id=ID_SEPARATOR;
	for (int i=0;i<_countof(g_StdCommands);i++)
		if (_wcsicmp(str,g_StdCommands[i].name)==0)
		{
			item.id=g_StdCommands[i].id;
			break;
		}
	if (item.id==ID_SEPARATOR)
	{
		item.id=ID_CUSTOM;
		item.command=str;
	}

	swprintf_s(text,L"%s.Icon",name);
	str=FindSetting(text);
	if (!str) return false;
	item.iconPath=str;

	swprintf_s(text,L"%s.IconDisabled",name);
	item.iconPathD=FindSetting(text);

	swprintf_s(text,L"%s.Tip",name);
	str=FindSetting(text);
	if (str)
	{
		if (str[0]=='$')
			item.tip=FindTranslation(str+1,NULL);
		else
			item.tip=str;
	}

	swprintf_s(text,L"%s.Name",name);
	str=FindSetting(text);
	if (str)
	{
		if (str[0]=='$')
			item.name=FindTranslation(str+1,NULL);
		else
			item.name=str;
	}

	return true;
}

void CBandWindow::ParseToolbar( DWORD enabled )
{
	m_Items.clear();
	const wchar_t *str=FindSetting("ToolbarItems");
	if (str)
	{
		// custom toolbar
		while (*str)
		{
			wchar_t token[256];
			str=GetToken(str,token,_countof(token),L", \t");
			StdToolbarItem item={ID_SEPARATOR};
			if (_wcsicmp(token,L"SEPARATOR")!=0)
			{
				if (!ParseToolbarItem(token,item))
					continue;
				if (item.id==ID_CUSTOM)
				{
					item.id=ID_CUSTOM+(int)m_Items.size();
					item.regName=token;
				}
			}
			m_Items.push_back(item);
		}
	}
	else
	{
		// standard toolbar
		for (int i=0;i<_countof(s_StdItems);i++)
		{
			if (enabled&(1<<s_StdItems[i].id))
			{
				m_Items.push_back(s_StdItems[i]);
				StdToolbarItem &item=m_Items[m_Items.size()-1];
				item.tip=FindTranslation(item.tipKey,item.tip);
			}
		}
	}
	if (m_Items.empty())
	{
		// make sure there is at least one button
		m_Items.push_back(s_StdItems[_countof(s_StdItems)-1]);
		m_Items[0].tip=FindTranslation(m_Items[0].tipKey,m_Items[0].tip);
	}
}

LRESULT CBandWindow::OnCreate( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled )
{
	CRegKey regSettings;
	if (regSettings.Open(HKEY_CURRENT_USER,L"Software\\IvoSoft\\ClassicExplorer")!=ERROR_SUCCESS)
		regSettings.Create(HKEY_CURRENT_USER,L"Software\\IvoSoft\\ClassicExplorer");

	DWORD BigButtons;
	if (regSettings.QueryDWORDValue(L"BigButtons",BigButtons)!=ERROR_SUCCESS)
		BigButtons=0;

	if (FindSetting("ToolbarItems"))
	{
		ParseToolbar(0);
	}
	else
	{
		DWORD ToolbarButtons=0;
		if (regSettings.QueryDWORDValue(L"ToolbarButtons",ToolbarButtons)!=ERROR_SUCCESS)
			ToolbarButtons=DEFAULT_BUTTONS|((ID_LAST-1)<<24);
		if (!(ToolbarButtons&0xFF000000)) ToolbarButtons|=0x07000002; // for backwards compatibility (when there were 7 buttons the button count was not saved)
		unsigned int mask1=(((2<<(ToolbarButtons>>24))-1)&~1); // bits to keep
		unsigned int mask2=(((2<<ID_LAST)-1)&~1)&~mask1; // bits to replace with defaults
		ToolbarButtons=(ToolbarButtons&mask1)|(DEFAULT_BUTTONS&mask2)|((ID_LAST-1)<<24);
		if ((ToolbarButtons&0xFFFFFF)==0)
			ToolbarButtons|=1<<DEFAULT_ONLY_BUTTON;
		ParseToolbar(ToolbarButtons);
	}

	bool bName=false;
	bool bList=FindSettingBool("ToolbarListMode",false);
	for (std::vector<StdToolbarItem>::const_iterator it=m_Items.begin();it!=m_Items.end();++it)
		if (it->name)
		{
			bName=true;
			break;
		}

	// create the toolbar
	if (bName && !bList)
		m_Toolbar=CreateWindow(TOOLBARCLASSNAME,L"",WS_CHILD|TBSTYLE_TOOLTIPS|TBSTYLE_FLAT|CCS_NODIVIDER|CCS_NOPARENTALIGN|CCS_NORESIZE,0,0,10,10,m_hWnd,(HMENU)101,g_Instance,NULL);
	else
		m_Toolbar=CreateWindow(TOOLBARCLASSNAME,L"",WS_CHILD|TBSTYLE_TOOLTIPS|TBSTYLE_FLAT|TBSTYLE_LIST|CCS_NODIVIDER|CCS_NOPARENTALIGN|CCS_NORESIZE,0,0,10,10,m_hWnd,(HMENU)101,g_Instance,NULL);

	m_Toolbar.SendMessage(TB_SETEXTENDEDSTYLE,0,TBSTYLE_EX_MIXEDBUTTONS);
	m_Toolbar.SendMessage(TB_BUTTONSTRUCTSIZE,sizeof(TBBUTTON));
	m_Toolbar.SendMessage(TB_SETMAXTEXTROWS,1);

	int iconSize=0;
	if (BigButtons)
	{
		const wchar_t *str=FindSetting("LargeIconSize");
		if (str) iconSize=_wtol(str);
	}
	else
	{
		const wchar_t *str=FindSetting("SmallIconSize");
		if (str) iconSize=_wtol(str);
	}

	if (iconSize==0)
	{
		// pick icon size based on the DPI setting
		HDC hdc=::GetDC(NULL);
		int dpi=GetDeviceCaps(hdc,LOGPIXELSY);
		::ReleaseDC(NULL,hdc);
		if (dpi>=120)
			iconSize=BigButtons?32:24;
		else
			iconSize=BigButtons?24:16;
	}
	else if (iconSize<8) iconSize=8;
	else if (iconSize>128) iconSize=128;

	m_Enabled=ImageList_Create(iconSize,iconSize,ILC_COLOR32|ILC_MASK|(IsLanguageRTL()?ILC_MIRROR:0),(int)m_Items.size(),2);
	m_Disabled=ImageList_Create(iconSize,iconSize,ILC_COLOR32|ILC_MASK|(IsLanguageRTL()?ILC_MIRROR:0),(int)m_Items.size(),2);

	HMODULE hShell32=GetModuleHandle(L"Shell32.dll");
	std::vector<HMODULE> modules;

	bool bSame=FindSettingBool("ToolbarSameSize",false);

	// create buttons
	std::vector<TBBUTTON> buttons(m_Items.size());
	for (int i=0;i<(int)m_Items.size();i++)
	{
		const StdToolbarItem &item=m_Items[i];
		TBBUTTON &button=buttons[i];

		button.idCommand=item.id;
		button.dwData=i;

		if (item.id==ID_SEPARATOR)
			button.fsStyle=BTNS_SEP;
		else
		{
			button.iBitmap=I_IMAGENONE;
			if (!item.iconPath || _wcsicmp(item.iconPath,L"NONE")!=0)
			{
				HICON hIcon=LoadIcon(iconSize,item.iconPath,item.icon,modules,hShell32);
				if (!hIcon)
					hIcon=(HICON)LoadImage(hShell32,MAKEINTRESOURCE(1),IMAGE_ICON,iconSize,iconSize,LR_DEFAULTCOLOR);
				if (hIcon)
				{
					button.iBitmap=ImageList_AddIcon(m_Enabled,hIcon);
					HICON hIcon2=item.iconPathD?LoadIcon(iconSize,item.iconPathD,0,modules,hShell32):NULL;
					if (!hIcon2)
						hIcon2=CreateDisabledIcon(hIcon,iconSize);
					int idx=ImageList_AddIcon(m_Disabled,hIcon2);
					ATLASSERT(button.iBitmap==idx);
					DestroyIcon(hIcon);
					DestroyIcon(hIcon2);
				}
			}

			button.fsState=(item.id!=ID_SETTINGS || FindSettingBool("EnableSettings",true))?TBSTATE_ENABLED:0;
			button.fsStyle=BTNS_BUTTON|BTNS_NOPREFIX;
			if (!bSame)
				button.fsStyle|=BTNS_AUTOSIZE;
			if (item.name)
			{
				button.fsStyle|=BTNS_SHOWTEXT;
				button.iString=(INT_PTR)item.name;
			}
		}
	}

	for (std::vector<HMODULE>::const_iterator it=modules.begin();it!=modules.end();++it)
		FreeLibrary(*it);

	// add buttons
	HIMAGELIST old=(HIMAGELIST)m_Toolbar.SendMessage(TB_SETIMAGELIST,0,(LPARAM)m_Enabled);
	if (old) ImageList_Destroy(old);
	old=(HIMAGELIST)m_Toolbar.SendMessage(TB_SETDISABLEDIMAGELIST,0,(LPARAM)m_Disabled);
	if (old) ImageList_Destroy(old);
	m_Toolbar.SendMessage(TB_ADDBUTTONS,buttons.size(),(LPARAM)&buttons[0]);
	SendMessage(WM_CLEAR);
	return 0;
}

LRESULT CBandWindow::OnDestroy( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled )
{
	ImageList_Destroy(m_Enabled);
	ImageList_Destroy(m_Disabled);
	bHandled=FALSE;
	return 0;
}

LRESULT CBandWindow::OnUpdateUI( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled )
{
	// update the state of the custom buttons based on the registry settings
	CRegKey regSettings;
	if (regSettings.Open(HKEY_CURRENT_USER,L"Software\\IvoSoft\\ClassicExplorer")==ERROR_SUCCESS)
	{
		for (std::vector<StdToolbarItem>::const_iterator it=m_Items.begin();it!=m_Items.end();++it)
		{
			if (!it->regName.empty())
			{
				DWORD val;
				if (regSettings.QueryDWORDValue(it->regName.c_str(),val)!=ERROR_SUCCESS)
					val=0;
				m_Toolbar.SendMessage(TB_ENABLEBUTTON,it->id,(val&1)?0:1);
				m_Toolbar.SendMessage(TB_CHECKBUTTON,it->id,(val&2)?1:0);
			}
		}
	}
	return 0;
}

// Go to the parent folder
LRESULT CBandWindow::OnNavigate( WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled )
{
	if (m_pBrowser)
	{
		UINT flags=(GetKeyState(VK_CONTROL)<0?SBSP_NEWBROWSER:SBSP_SAMEBROWSER);
		if (wID==ID_GOUP)
			m_pBrowser->BrowseObject(NULL,flags|SBSP_PARENT);
		if (wID==ID_GOBACK)
			m_pBrowser->BrowseObject(NULL,flags|SBSP_NAVIGATEBACK);
		if (wID==ID_GOFORWARD)
			m_pBrowser->BrowseObject(NULL,flags|SBSP_NAVIGATEFORWARD);
	}

	return TRUE;
}

void CBandWindow::SendShellTabCommand( int command )
{
	// sends a command to the ShellTabWindowClass window
	for (CWindow parent=GetParent();parent.m_hWnd;parent=parent.GetParent())
	{
		// find a parent window with class ShellTabWindowClass
		wchar_t name[256];
		GetClassName(parent.m_hWnd,name,_countof(name));
		if (_wcsicmp(name,L"ShellTabWindowClass")==0)
		{
			parent.SendMessage(WM_COMMAND,command);
			break;
		}
	}
}

// Executes a cut/copy/paste/delete command
LRESULT CBandWindow::OnToolbarCommand( WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled )
{
	if (wID>=ID_CUSTOM)
	{
		int idx=wID-ID_CUSTOM;
		wchar_t buf[2048];
		wcscpy_s(buf,m_Items[idx].command);
		DoEnvironmentSubst(buf,_countof(buf));
		wchar_t *pBuf=buf;
		bool bArg1=wcsstr(buf,L"%1")!=NULL;
		bool bArg2=wcsstr(buf,L"%2")!=NULL;
		wchar_t path[_MAX_PATH];
		wchar_t file[_MAX_PATH];
		path[0]=file[0]=0;

		CComPtr<IShellView> pView;
		if (SUCCEEDED(m_pBrowser->QueryActiveShellView(&pView)))
		{
			CComPtr<IPersistFolder2> pFolder;
			LPITEMIDLIST pidl;
			CComQIPtr<IFolderView> pView2=pView;
			if (pView2 && SUCCEEDED(pView2->GetFolder(IID_IPersistFolder2,(void**)&pFolder)) && SUCCEEDED(pFolder->GetCurFolder(&pidl)))
			{
				// get current path
				SHGetPathFromIDList(pidl,path);
				if (bArg2)
				{
					CComPtr<IEnumIDList> pEnum;
					int count;
					// if only one file is selected get the file name (%2)
					if (SUCCEEDED(pView2->ItemCount(SVGIO_SELECTION,&count)) && count==1 && SUCCEEDED(pView2->Items(SVGIO_SELECTION,IID_IEnumIDList,(void**)&pEnum)) && pEnum)
					{
						PITEMID_CHILD child;
						if (pEnum->Next(1,&child,NULL)==S_OK)
						{
							LPITEMIDLIST full=ILCombine(pidl,child);
							SHGetPathFromIDList(full,file);
							ILFree(child);
							ILFree(full);
						}
					}
				}
				ILFree(pidl);
			}
		}

		if (bArg1 || bArg2)
		{
			// expand environment variables, %1, %2
			DWORD_PTR args[100]={(DWORD_PTR)path,(DWORD_PTR)file};
			FormatMessage(FORMAT_MESSAGE_ALLOCATE_BUFFER|FORMAT_MESSAGE_ARGUMENT_ARRAY|FORMAT_MESSAGE_FROM_STRING,buf,0,0,(LPWSTR)&pBuf,0,(va_list*)args);
		}

		wchar_t exe[_MAX_PATH];
		const wchar_t *params=GetToken(pBuf,exe,_countof(exe),L" ");
		ShellExecute(NULL,NULL,exe,params,path,SW_SHOWNORMAL);
		if (pBuf!=buf)
			LocalFree(pBuf);
		return TRUE;
	}
	// check if the focus is on the tree side or on the list side
	CWindow focus=GetFocus();
	wchar_t name[256];
	GetClassName(focus,name,_countof(name));
	CWindow parent=focus.GetParent();
	if (_wcsicmp(name,WC_TREEVIEW)==0)
	{
		// send these commands to the parent of the tree view
		if (wID==ID_CUT)
			parent.SendMessage(WM_COMMAND,41025);
		if (wID==ID_COPY)
			parent.SendMessage(WM_COMMAND,41026);
		if (wID==ID_PASTE)
			parent.SendMessage(WM_COMMAND,41027);
		if (wID==ID_DELETE)
			parent.SendMessage(WM_COMMAND,40995);
		if (wID==ID_PROPERTIES)
			ShowTreeProperties(focus.m_hWnd);
	}
	else
	{
		GetClassName(parent,name,_countof(name));
		if (_wcsicmp(name,L"SHELLDLL_DefView")==0)
		{
			// send these commands to the SHELLDLL_DefView window
			if (wID==ID_CUT)
			{
				parent.SendMessage(WM_COMMAND,28696);
				focus.InvalidateRect(NULL);
			}
			if (wID==ID_COPY)
				parent.SendMessage(WM_COMMAND,28697);
			if (wID==ID_PASTE)
				parent.SendMessage(WM_COMMAND,28698);
			if (wID==ID_DELETE)
				parent.SendMessage(WM_COMMAND,28689);
			if (wID==ID_PROPERTIES)
				parent.SendMessage(WM_COMMAND,28691);
			if (wID==ID_COPYTO)
				parent.SendMessage(WM_COMMAND,28702);
			if (wID==ID_MOVETO)
				parent.SendMessage(WM_COMMAND,28703);
		}
	}

	if (wID==ID_UNDO)
		SendShellTabCommand(28699);
	if (wID==ID_REDO)
		SendShellTabCommand(28704);
	if (wID==ID_SELECTALL)
		SendShellTabCommand(28705);
	if (wID==ID_INVERT)
		SendShellTabCommand(28706);
	if (wID==ID_REFRESH)
		SendShellTabCommand(41504);

	return TRUE;
}

LRESULT CBandWindow::OnEmail( WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled )
{
	const IID CLSID_SendMail={0x9E56BE60,0xC50F,0x11CF,{0x9A,0x2C,0x00,0xA0,0xC9,0x0A,0x90,0xCE}};

	CComPtr<IShellView> pView;
	if (FAILED(m_pBrowser->QueryActiveShellView(&pView))) return TRUE;

	// check if there is anything selected
	CComQIPtr<IFolderView> pView2=pView;
	int count;
	if (pView2 && SUCCEEDED(pView2->ItemCount(SVGIO_SELECTION,&count)) && count==0) return TRUE;

	// get the data object
	CComPtr<IDataObject> pDataObj;
	if (FAILED(pView->GetItemObject(SVGIO_SELECTION,IID_IDataObject,(void**)&pDataObj))) return TRUE;
	CComQIPtr<IAsyncOperation> pAsync=pDataObj;
	if (pAsync)
		pAsync->SetAsyncMode(FALSE);

	// drop into the SendMail handler
	CComPtr<IDropTarget> pDropTarget;
	if (SUCCEEDED(CoCreateInstance(CLSID_SendMail,NULL,CLSCTX_ALL,IID_IDropTarget,(void **)&pDropTarget)))
	{
		POINTL pt={0,0};
		DWORD dwEffect=0;
		pDropTarget->DragEnter(pDataObj,MK_LBUTTON,pt,&dwEffect);
		pDropTarget->Drop(pDataObj,0,pt,&dwEffect);
	}
	return TRUE;
}

// Show the settings dialog
LRESULT CBandWindow::OnSettings( WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled )
{
#ifndef BUILD_SETUP
	if (GetKeyState(VK_SHIFT)<0)
		*(int*)0=0; // force a crash if Shift is pressed. Makes it easy to restart explorer.exe
#endif
	ShowSettings(m_hWnd);
	return TRUE;
}

LRESULT CBandWindow::OnRClick( int idCtrl, LPNMHDR pnmh, BOOL& bHandled )
{
	NMMOUSE *pInfo=(NMMOUSE*)pnmh;
	POINT pt=pInfo->pt;
	{
		RECT rc;
		int count=(int)m_Toolbar.SendMessage(TB_BUTTONCOUNT);
		m_Toolbar.SendMessage(TB_GETITEMRECT,count-1,(LPARAM)&rc);
		if (pt.x>rc.right)
			return 0;
	}
	m_Toolbar.ClientToScreen(&pt);
	ShowSettingsMenu(m_hWnd,pt.x,pt.y);
	return 1;
}

LRESULT CBandWindow::OnGetInfoTip( int idCtrl, LPNMHDR pnmh, BOOL& bHandled )
{
	NMTBGETINFOTIP *pTip=(NMTBGETINFOTIP*)pnmh;
	const StdToolbarItem &item=m_Items[pTip->lParam];
	if (item.tip)
	{
		// show the tip for the standard item
		wcscpy_s(pTip->pszText,pTip->cchTextMax,item.tip);
	}
	return 0;
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
	m_Toolbar.SendMessage(TB_ENABLEBUTTON,CBandWindow::ID_GOUP,bDesktop?0:1);
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
	IObjectWithSiteImpl<CExplorerBand>::SetSite(pUnkSite);

	if (m_BandWindow.IsWindow())
		m_BandWindow.DestroyWindow();
	m_BandWindow.SetBrowser(NULL);
	if (m_bSubclassedRebar)
		RemoveWindowSubclass(GetParent(m_BandWindow.GetToolbar()),RebarSubclassProc,(UINT_PTR)this);
	m_bSubclassedRebar=false;

	if (m_pWebBrowser && m_dwEventCookie!=0xFEFEFEFE)
		DispEventUnadvise(m_pWebBrowser,&DIID_DWebBrowserEvents2);
	m_pWebBrowser=NULL;

	//If punkSite is not NULL, a new site is being set.
	if (pUnkSite)
	{
extern void ReadIniFile( bool bStartup );
		ReadIniFile(false);

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

STDMETHODIMP CExplorerBand::OnNavigateComplete( IDispatch *pDisp, VARIANT *URL )
{
	// this is called when the current folder changes. disable the Up button if this is the desktop folder
	m_BandWindow.UpdateToolbar();
	return S_OK;
}

STDMETHODIMP CExplorerBand::OnCommandStateChange( long Command, VARIANT_BOOL Enable )
{
	if (Command==CSC_NAVIGATEFORWARD)
	{
		SendMessage(m_BandWindow.GetToolbar(),TB_ENABLEBUTTON,CBandWindow::ID_GOFORWARD,Enable?1:0);
	}
	if (Command==CSC_NAVIGATEBACK)
	{
		SendMessage(m_BandWindow.GetToolbar(),TB_ENABLEBUTTON,CBandWindow::ID_GOBACK,Enable?1:0);
	}
	return S_OK;
}

STDMETHODIMP CExplorerBand::OnQuit( void )
{
	if (m_pWebBrowser && m_dwEventCookie!=0xFEFEFEFE) // ATL's event cookie is 0xFEFEFEFE, when the sink is not advised
		return DispEventUnadvise(m_pWebBrowser,&DIID_DWebBrowserEvents2);
	return S_OK;
}
