// Classic Shell (c) 2009-2010, Ivo Beltchev
// The sources for Classic Shell are distributed under the MIT open source license

// ExplorerBand.cpp : Implementation of CExplorerBand

#include "stdafx.h"
#include "ExplorerBand.h"
#include "resource.h"
#include "ExplorerBHO.h"
#include "GlobalSettings.h"
#include "TranslationSettings.h"
#include <vector>

const int DEFAULT_BUTTONS=0x1FE; // buttons visible by default
const int DEFAULT_ONLY_BUTTON=CBandWindow::ID_SETTINGS; // use this button when all buttons are hidden (there must be at least one visible button)

// Dialog proc for the Settings dialog. Edits the settings and saves them to the registry
INT_PTR CALLBACK SettingsDlgProc( HWND hwndDlg, UINT uMsg, WPARAM wParam, LPARAM lParam )
{
	if (uMsg==WM_INITDIALOG)
	{
		void *pRes=NULL;
		HRSRC hResInfo=FindResource(g_Instance,MAKEINTRESOURCE(VS_VERSION_INFO),RT_VERSION);
		if (hResInfo)
		{
			HGLOBAL hRes=LoadResource(g_Instance,hResInfo);
			pRes=LockResource(hRes);
		}
		wchar_t title[100];
		if (pRes)
		{
			VS_FIXEDFILEINFO *pVer=(VS_FIXEDFILEINFO*)((char*)pRes+40);
			swprintf_s(title,L"Settings for Classic Explorer %d.%d.%d",HIWORD(pVer->dwProductVersionMS),LOWORD(pVer->dwProductVersionMS),HIWORD(pVer->dwProductVersionLS));
		}
		else
			swprintf_s(title,L"Settings for Classic Explorer");
		SetWindowText(hwndDlg,title);

		HICON icon=(HICON)LoadImage(g_Instance,MAKEINTRESOURCE(IDI_APPICON),IMAGE_ICON,GetSystemMetrics(SM_CXICON),GetSystemMetrics(SM_CYICON),LR_DEFAULTCOLOR);
		SendMessage(hwndDlg,WM_SETICON,ICON_BIG,(LPARAM)icon);
		icon=(HICON)LoadImage(g_Instance,MAKEINTRESOURCE(IDI_APPICON),IMAGE_ICON,GetSystemMetrics(SM_CXSMICON),GetSystemMetrics(SM_CYSMICON),LR_DEFAULTCOLOR);
		SendMessage(hwndDlg,WM_SETICON,ICON_SMALL,(LPARAM)icon);

		CRegKey regSettings;
		if (regSettings.Open(HKEY_CURRENT_USER,L"Software\\IvoSoft\\ClassicExplorer")!=ERROR_SUCCESS)
			regSettings.Create(HKEY_CURRENT_USER,L"Software\\IvoSoft\\ClassicExplorer");

		DWORD EnableCopyUI, FreeSpace, FoldersSettings, BigButtons, ToolbarButtons;
		if (regSettings.QueryDWORDValue(L"EnableCopyUI",EnableCopyUI)!=ERROR_SUCCESS)
			EnableCopyUI=1;
		if (regSettings.QueryDWORDValue(L"FreeSpace",FreeSpace)!=ERROR_SUCCESS)
			FreeSpace=(LOWORD(GetVersion())==0x0106)?CExplorerBHO::SPACE_SHOW:0;
		if (regSettings.QueryDWORDValue(L"FoldersSettings",FoldersSettings)!=ERROR_SUCCESS)
			FoldersSettings=CExplorerBHO::FOLDERS_DEFAULT;
		if (regSettings.QueryDWORDValue(L"BigButtons",BigButtons)!=ERROR_SUCCESS)
			BigButtons=0;
		if (regSettings.QueryDWORDValue(L"ToolbarButtons",ToolbarButtons)!=ERROR_SUCCESS)
			ToolbarButtons=DEFAULT_BUTTONS|((CBandWindow::ID_LAST-1)<<24);

		if (!(ToolbarButtons&0xFF000000)) ToolbarButtons|=0x07000002; // for backwards compatibility (when there were 7 buttons the the button count was not saved)
		unsigned int mask1=(((2<<(ToolbarButtons>>24))-1)&~1); // bits to keep
		unsigned int mask2=(((2<<CBandWindow::ID_LAST)-1)&~1)&~mask1; // bits to replace with defaults
		ToolbarButtons=(ToolbarButtons&mask1)|(DEFAULT_BUTTONS&mask2)|((CBandWindow::ID_LAST-1)<<24);
		if ((ToolbarButtons&0xFFFFFF)==0)
			ToolbarButtons|=1<<DEFAULT_ONLY_BUTTON;

		RECT rc1,rc2;
		GetWindowRect(hwndDlg,&rc1);
		GetWindowRect(GetParent(hwndDlg),&rc2);
		OffsetRect(&rc1,(rc2.left+rc2.right)/2-(rc1.left+rc1.right)/2,(rc2.top+rc2.bottom)/2-(rc1.top+rc1.bottom)/2);
		if (rc1.top<rc2.top) OffsetRect(&rc1,0,rc2.top-rc1.top);
		SetWindowPos(hwndDlg,NULL,rc1.left,rc1.top,rc1.right-rc1.left,rc1.bottom-rc1.top,SWP_NOZORDER);
		SendMessage(hwndDlg,DM_REPOSITION,0,0);

		CheckDlgButton(hwndDlg,IDC_CHECKCOPY,(EnableCopyUI==1 || EnableCopyUI==2)?BST_CHECKED:BST_UNCHECKED);
		CheckDlgButton(hwndDlg,IDC_CHECKCOPYFOLDER,(EnableCopyUI&1)?BST_CHECKED:BST_UNCHECKED);
		CheckDlgButton(hwndDlg,IDC_CHECKSIZE,FreeSpace?BST_CHECKED:BST_UNCHECKED);
		CheckDlgButton(hwndDlg,IDC_CHECKBHO,(FoldersSettings&CExplorerBHO::FOLDERS_ALTENTER)?BST_CHECKED:BST_UNCHECKED);
		CheckDlgButton(hwndDlg,IDC_CHECKAUTO,(FoldersSettings&CExplorerBHO::FOLDERS_AUTONAVIGATE)?BST_CHECKED:BST_UNCHECKED);
		CheckDlgButton(hwndDlg,IDC_CHECKXPSTYLE,(FoldersSettings&CExplorerBHO::FOLDERS_CLASSIC)?BST_CHECKED:BST_UNCHECKED);
		CheckDlgButton(hwndDlg,IDC_CHECKSIMPLE,(FoldersSettings&CExplorerBHO::FOLDERS_SIMPLE)?BST_CHECKED:BST_UNCHECKED);
		CheckDlgButton(hwndDlg,IDC_CHECKNOFADE,(FoldersSettings&CExplorerBHO::FOLDERS_NOFADE)?BST_CHECKED:BST_UNCHECKED);
		CheckDlgButton(hwndDlg,IDC_CHECKBIG,BigButtons?BST_CHECKED:BST_UNCHECKED);
		CheckDlgButton(hwndDlg,IDC_CHECK1,(ToolbarButtons&(1<<CBandWindow::ID_GOUP))?BST_CHECKED:BST_UNCHECKED);
		CheckDlgButton(hwndDlg,IDC_CHECK2,(ToolbarButtons&(1<<CBandWindow::ID_CUT))?BST_CHECKED:BST_UNCHECKED);
		CheckDlgButton(hwndDlg,IDC_CHECK3,(ToolbarButtons&(1<<CBandWindow::ID_COPY))?BST_CHECKED:BST_UNCHECKED);
		CheckDlgButton(hwndDlg,IDC_CHECK4,(ToolbarButtons&(1<<CBandWindow::ID_PASTE))?BST_CHECKED:BST_UNCHECKED);
		CheckDlgButton(hwndDlg,IDC_CHECK5,(ToolbarButtons&(1<<CBandWindow::ID_DELETE))?BST_CHECKED:BST_UNCHECKED);
		CheckDlgButton(hwndDlg,IDC_CHECK6,(ToolbarButtons&(1<<CBandWindow::ID_PROPERTIES))?BST_CHECKED:BST_UNCHECKED);
		CheckDlgButton(hwndDlg,IDC_CHECK7,(ToolbarButtons&(1<<CBandWindow::ID_EMAIL))?BST_CHECKED:BST_UNCHECKED);
		CheckDlgButton(hwndDlg,IDC_CHECK8,(ToolbarButtons&(1<<CBandWindow::ID_SETTINGS))?BST_CHECKED:BST_UNCHECKED);
	}
	if (uMsg==WM_INITDIALOG || (uMsg==WM_COMMAND && (wParam==IDC_CHECKXPSTYLE || wParam==IDC_CHECKSIMPLE)))
	{
		BOOL bXPStyle=IsDlgButtonChecked(hwndDlg,IDC_CHECKXPSTYLE)==BST_CHECKED;
		BOOL bSimple=IsDlgButtonChecked(hwndDlg,IDC_CHECKSIMPLE)==BST_CHECKED;
		EnableWindow(GetDlgItem(hwndDlg,IDC_CHECKSIMPLE),bXPStyle);
		EnableWindow(GetDlgItem(hwndDlg,IDC_CHECKNOFADE),!bXPStyle || bSimple);
		if (uMsg==WM_COMMAND)
			return TRUE;
	}
	if (uMsg==WM_INITDIALOG || (uMsg==WM_COMMAND && wParam==IDC_CHECKCOPY))
	{
		EnableWindow(GetDlgItem(hwndDlg,IDC_CHECKCOPYFOLDER),IsDlgButtonChecked(hwndDlg,IDC_CHECKCOPY)==BST_CHECKED);
		if (uMsg==WM_COMMAND)
			return TRUE;
	}
	if (uMsg==WM_INITDIALOG)
		return TRUE;
	if (uMsg==WM_COMMAND && wParam==IDOK)
	{
		CRegKey regSettings;
		if (regSettings.Open(HKEY_CURRENT_USER,L"Software\\IvoSoft\\ClassicExplorer")!=ERROR_SUCCESS)
			regSettings.Create(HKEY_CURRENT_USER,L"Software\\IvoSoft\\ClassicExplorer");

		DWORD EnableCopyUI, FreeSpace, FoldersSettings, BigButtons, ToolbarButtons;
		if (regSettings.QueryDWORDValue(L"EnableCopyUI",EnableCopyUI)!=ERROR_SUCCESS)
			EnableCopyUI=1;
		if (regSettings.QueryDWORDValue(L"FreeSpace",FreeSpace)!=ERROR_SUCCESS)
			FreeSpace=(LOWORD(GetVersion())==0x0106)?CExplorerBHO::SPACE_SHOW:0;
		if (regSettings.QueryDWORDValue(L"FoldersSettings",FoldersSettings)!=ERROR_SUCCESS)
			FoldersSettings=CExplorerBHO::FOLDERS_DEFAULT;
		if (regSettings.QueryDWORDValue(L"BigButtons",BigButtons)!=ERROR_SUCCESS)
			BigButtons=0;
		if (regSettings.QueryDWORDValue(L"ToolbarButtons",ToolbarButtons)!=ERROR_SUCCESS)
			ToolbarButtons=DEFAULT_BUTTONS|((CBandWindow::ID_LAST-1)<<24);
		DWORD ToolbarButtons0=ToolbarButtons;
		if (!(ToolbarButtons&0xFF000000)) ToolbarButtons|=0x07000002; // for backwards compatibility (when there were 7 buttons the the button count was not saved)
		unsigned int mask1=(((2<<(ToolbarButtons>>24))-1)&~1); // bits to keep
		unsigned int mask2=(((2<<CBandWindow::ID_LAST)-1)&~1)&~mask1; // bits to replace with defaults
		ToolbarButtons=(ToolbarButtons&mask1)|(DEFAULT_BUTTONS&mask2)|((CBandWindow::ID_LAST-1)<<24);
		if ((ToolbarButtons&0xFFFFFF)==0)
			ToolbarButtons|=1<<DEFAULT_ONLY_BUTTON;

		DWORD EnableCopyUI2=(IsDlgButtonChecked(hwndDlg,IDC_CHECKCOPY)==BST_CHECKED)?2:0;
		if (IsDlgButtonChecked(hwndDlg,IDC_CHECKCOPYFOLDER)==BST_CHECKED)
			EnableCopyUI2=EnableCopyUI2?1:3;
		DWORD FreeSpace2=(IsDlgButtonChecked(hwndDlg,IDC_CHECKSIZE)==BST_CHECKED)?CExplorerBHO::SPACE_SHOW:0;
		DWORD FoldersSettings2=0;
		if (IsDlgButtonChecked(hwndDlg,IDC_CHECKBHO)==BST_CHECKED)
			FoldersSettings2|=CExplorerBHO::FOLDERS_ALTENTER;
		if (IsDlgButtonChecked(hwndDlg,IDC_CHECKAUTO)==BST_CHECKED)
			FoldersSettings2|=CExplorerBHO::FOLDERS_AUTONAVIGATE;
		if (IsDlgButtonChecked(hwndDlg,IDC_CHECKXPSTYLE)==BST_CHECKED)
		{
			FoldersSettings2|=CExplorerBHO::FOLDERS_CLASSIC;
			if (IsDlgButtonChecked(hwndDlg,IDC_CHECKSIMPLE)==BST_CHECKED)
				FoldersSettings2|=CExplorerBHO::FOLDERS_SIMPLE;
		}
		if (IsDlgButtonChecked(hwndDlg,IDC_CHECKNOFADE)==BST_CHECKED)
			FoldersSettings2|=CExplorerBHO::FOLDERS_NOFADE;

		DWORD BigButtons2=(IsDlgButtonChecked(hwndDlg,IDC_CHECKBIG)==BST_CHECKED)?1:0;
		DWORD ToolbarButtons2=0;
		ToolbarButtons2|=(IsDlgButtonChecked(hwndDlg,IDC_CHECK1)==BST_CHECKED)?(1<<CBandWindow::ID_GOUP):0;
		ToolbarButtons2|=(IsDlgButtonChecked(hwndDlg,IDC_CHECK2)==BST_CHECKED)?(1<<CBandWindow::ID_CUT):0;
		ToolbarButtons2|=(IsDlgButtonChecked(hwndDlg,IDC_CHECK3)==BST_CHECKED)?(1<<CBandWindow::ID_COPY):0;
		ToolbarButtons2|=(IsDlgButtonChecked(hwndDlg,IDC_CHECK4)==BST_CHECKED)?(1<<CBandWindow::ID_PASTE):0;
		ToolbarButtons2|=(IsDlgButtonChecked(hwndDlg,IDC_CHECK5)==BST_CHECKED)?(1<<CBandWindow::ID_DELETE):0;
		ToolbarButtons2|=(IsDlgButtonChecked(hwndDlg,IDC_CHECK6)==BST_CHECKED)?(1<<CBandWindow::ID_PROPERTIES):0;
		ToolbarButtons2|=(IsDlgButtonChecked(hwndDlg,IDC_CHECK7)==BST_CHECKED)?(1<<CBandWindow::ID_EMAIL):0;
		ToolbarButtons2|=(IsDlgButtonChecked(hwndDlg,IDC_CHECK8)==BST_CHECKED)?(1<<CBandWindow::ID_SETTINGS):0;
		ToolbarButtons2|=((CBandWindow::ID_LAST-1)<<24);
		if ((ToolbarButtons2&0xFFFFFF)==0)
			ToolbarButtons2|=1<<DEFAULT_ONLY_BUTTON;

		int res=0;
		if (EnableCopyUI!=EnableCopyUI2)
		{
extern bool g_bHookCopyThreads;
			if (!g_bHookCopyThreads && (EnableCopyUI2==1 || EnableCopyUI2==2))
				res|=2;
			regSettings.SetDWORDValue(L"EnableCopyUI",EnableCopyUI2);
		}
		if (FreeSpace!=FreeSpace2)
		{
			regSettings.SetDWORDValue(L"FreeSpace",FreeSpace2);
			res|=1;
		}
		if (FoldersSettings!=FoldersSettings2)
		{
			regSettings.SetDWORDValue(L"FoldersSettings",FoldersSettings2);
			if ((FoldersSettings^FoldersSettings2)!=CExplorerBHO::FOLDERS_ALTENTER)
				res|=1;
		}
		if (BigButtons!=BigButtons2)
		{
			regSettings.SetDWORDValue(L"BigButtons",BigButtons2);
			res|=1;
		}
		if (ToolbarButtons0!=ToolbarButtons2)
		{
			regSettings.SetDWORDValue(L"ToolbarButtons",ToolbarButtons2);
			if (ToolbarButtons!=ToolbarButtons2)
				res|=1;
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

static HICON LoadIcon( int iconSize, const char *setting, int index, std::vector<HMODULE> &modules, HMODULE hShell32 )
{
	const wchar_t *str=FindSetting(setting);
	if (!str)
		return (HICON)LoadImage(hShell32,MAKEINTRESOURCE(index),IMAGE_ICON,iconSize,iconSize,LR_DEFAULTCOLOR);
	wchar_t text[1024];
	wcscpy_s(text,str);
	DoEnvironmentSubst(text,_countof(text));
	wchar_t *c=wcsrchr(text,',');
	if (c)
	{
		// resource file
		*c=0;
		index=_wtol(c+1);
		HMODULE hMod=GetModuleHandle(PathFindFileName(text));
		if (!hMod)
		{
			hMod=LoadLibraryEx(text,NULL,LOAD_LIBRARY_AS_DATAFILE|LOAD_LIBRARY_AS_IMAGE_RESOURCE);
			if (!hMod) return NULL;
			modules.push_back(hMod);
		}
		return (HICON)LoadImage(hMod,MAKEINTRESOURCE(index),IMAGE_ICON,iconSize,iconSize,LR_DEFAULTCOLOR);
	}
	else
	{
		return (HICON)LoadImage(NULL,text,IMAGE_ICON,iconSize,iconSize,LR_DEFAULTCOLOR|LR_LOADFROMFILE);
	}
}

LRESULT CBandWindow::OnCreate( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled )
{
	// create the toolbar
	m_Toolbar=CreateWindow(TOOLBARCLASSNAME,L"",WS_CHILD|TBSTYLE_TOOLTIPS|TBSTYLE_FLAT|TBSTYLE_LIST|CCS_NODIVIDER|CCS_NOPARENTALIGN|CCS_NORESIZE,0,0,10,10,m_hWnd,(HMENU)101,g_Instance,NULL);

	m_Toolbar.SendMessage(TB_SETEXTENDEDSTYLE,0,TBSTYLE_EX_MIXEDBUTTONS);
	m_Toolbar.SendMessage(TB_BUTTONSTRUCTSIZE,sizeof(TBBUTTON));
	m_Toolbar.SendMessage(TB_SETMAXTEXTROWS,1);

	CRegKey regSettings;
	if (regSettings.Open(HKEY_CURRENT_USER,L"Software\\IvoSoft\\ClassicExplorer")!=ERROR_SUCCESS)
		regSettings.Create(HKEY_CURRENT_USER,L"Software\\IvoSoft\\ClassicExplorer");

	DWORD BigButtons, ToolbarButtons;
	if (regSettings.QueryDWORDValue(L"BigButtons",BigButtons)!=ERROR_SUCCESS)
		BigButtons=0;
	if (regSettings.QueryDWORDValue(L"ToolbarButtons",ToolbarButtons)!=ERROR_SUCCESS)
		ToolbarButtons=DEFAULT_BUTTONS|((ID_LAST-1)<<24);
	if (!(ToolbarButtons&0xFF000000)) ToolbarButtons|=0x07000002; // for backwards compatibility (when there were 7 buttons the the button count was not saved)
	unsigned int mask1=(((2<<(ToolbarButtons>>24))-1)&~1); // bits to keep
	unsigned int mask2=(((2<<ID_LAST)-1)&~1)&~mask1; // bits to replace with defaults
	ToolbarButtons=(ToolbarButtons&mask1)|(DEFAULT_BUTTONS&mask2)|((ID_LAST-1)<<24);
	if ((ToolbarButtons&0xFFFFFF)==0)
		ToolbarButtons|=1<<DEFAULT_ONLY_BUTTON;

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

	m_Enabled=ImageList_Create(iconSize,iconSize,ILC_COLOR32|ILC_MASK|(IsLanguageRTL()?ILC_MIRROR:0),0,2);

	TBBUTTON buttons[]={
		{46,ID_GOUP,TBSTATE_ENABLED,BTNS_BUTTON|BTNS_AUTOSIZE,{0},(DWORD_PTR)"IconUp",(INT_PTR)FindTranslation("Toolbar.GoUp",L"Up One Level")},
		{16762,ID_CUT,TBSTATE_ENABLED,BTNS_BUTTON|BTNS_AUTOSIZE,{0},(DWORD_PTR)"IconCut",(INT_PTR)FindTranslation("Toolbar.Cut",L"Cut")},
		{243,ID_COPY,TBSTATE_ENABLED,BTNS_BUTTON|BTNS_AUTOSIZE,{0},(DWORD_PTR)"IconCopy",(INT_PTR)FindTranslation("Toolbar.Copy",L"Copy")},
		{16763,ID_PASTE,TBSTATE_ENABLED,BTNS_BUTTON|BTNS_AUTOSIZE,{0},(DWORD_PTR)"IconPaste",(INT_PTR)FindTranslation("Toolbar.Paste",L"Paste")},
		{240,ID_DELETE,TBSTATE_ENABLED,BTNS_BUTTON|BTNS_AUTOSIZE,{0},(DWORD_PTR)"IconDelete",(INT_PTR)FindTranslation("Toolbar.Delete",L"Delete")},
		{253,ID_PROPERTIES,TBSTATE_ENABLED,BTNS_BUTTON|BTNS_AUTOSIZE,{0},(DWORD_PTR)"IconProperties",(INT_PTR)FindTranslation("Toolbar.Properties",L"Properties")},
		{265,ID_EMAIL,TBSTATE_ENABLED,BTNS_BUTTON|BTNS_AUTOSIZE,{0},(DWORD_PTR)"IconEmail",(INT_PTR)FindTranslation("Toolbar.Email",L"E-mail the selected items")},
		{IDI_APPICON,ID_SETTINGS,TBSTATE_ENABLED,BTNS_BUTTON|BTNS_AUTOSIZE,{0},(DWORD_PTR)"IconSettings",(INT_PTR)FindTranslation("Toolbar.Settings",L"Classic Explorer Settings")},
	};

	// load icons from Shell32.dll
	HMODULE hShell32=GetModuleHandle(L"Shell32.dll");
	std::vector<HMODULE> modules;

	int count=0; // count up until the last visible button and don't create the rest (we use the last button to get the total toolbar size)
	for (int i=0;i<_countof(buttons);i++)
	{
		if (!(ToolbarButtons&(1<<buttons[i].idCommand)))
		{
			buttons[i].fsState|=TBSTATE_HIDDEN;
			buttons[i].iBitmap=I_IMAGENONE;
			continue;
		}
		count=i+1;
		HICON hIcon=LoadIcon(iconSize,(const char*)buttons[i].dwData,buttons[i].iBitmap,modules,(buttons[i].idCommand==ID_SETTINGS?g_Instance:hShell32));
		if (!hIcon)
			hIcon=(HICON)LoadImage(hShell32,MAKEINTRESOURCE(1),IMAGE_ICON,iconSize,iconSize,LR_DEFAULTCOLOR);
		if (hIcon)
		{
			buttons[i].iBitmap=ImageList_AddIcon(m_Enabled,hIcon);
			DestroyIcon(hIcon);
		}
		else
			buttons[i].iBitmap=I_IMAGENONE;
	}

	for (std::vector<HMODULE>::const_iterator it=modules.begin();it!=modules.end();++it)
		FreeLibrary(*it);

	// add buttons
	HIMAGELIST old=(HIMAGELIST)m_Toolbar.SendMessage(TB_SETIMAGELIST,0,(LPARAM)m_Enabled);
	if (old) ImageList_Destroy(old);
	m_Toolbar.SendMessage(TB_ADDBUTTONS,count,(LPARAM)buttons);
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

// Executes a cut/copy/paste/delete command
LRESULT CBandWindow::OnFileOperation( WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled )
{
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

		return TRUE;
	}

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
	}

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
	INT_PTR res=DialogBox(g_Instance,MAKEINTRESOURCE(IDD_SETTINGS),m_hWnd,SettingsDlgProc);
	if (res&2)
		MessageBox(L"You need to log off and back on for the new settings to take effect.",L"Classic Explorer");
	else if (res&1)
		MessageBox(L"The new settings will take effect the next time you open an Explorer window.",L"Classic Explorer");
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
	HMENU menu=CreatePopupMenu();
	AppendMenu(menu,MF_STRING,ID_SETTINGS,FindTranslation("Toolbar.Settings",L"Classic Explorer Settings"));
	HBITMAP shellBmp=NULL;
	int size=16;
	BITMAPINFO bi={0};
	bi.bmiHeader.biSize=sizeof(BITMAPINFOHEADER);
	bi.bmiHeader.biWidth=bi.bmiHeader.biHeight=size;
	bi.bmiHeader.biPlanes=1;
	bi.bmiHeader.biBitCount=32;
	HDC hdc=CreateCompatibleDC(NULL);
	RECT rc={0,0,size,size};
	HICON icon=(HICON)LoadImage(g_Instance,MAKEINTRESOURCE(IDI_APPICON),IMAGE_ICON,size,size,LR_DEFAULTCOLOR);
	if (icon)
	{
		shellBmp=CreateDIBSection(hdc,&bi,DIB_RGB_COLORS,NULL,NULL,0);
		HGDIOBJ bmp0=SelectObject(hdc,shellBmp);
		FillRect(hdc,&rc,(HBRUSH)GetStockObject(BLACK_BRUSH));
		DrawIconEx(hdc,0,0,icon,size,size,0,NULL,DI_NORMAL);
		SelectObject(hdc,bmp0);
		DeleteObject(icon);

		MENUITEMINFO mii={sizeof(mii)};
		mii.fMask=MIIM_BITMAP;
		mii.hbmpItem=shellBmp;
		SetMenuItemInfo(menu,ID_SETTINGS,FALSE,&mii);
	}
	DeleteDC(hdc);

	DWORD pos=GetMessagePos();
	TrackPopupMenu(menu,0,pt.x,pt.y,0,m_hWnd,NULL);
	DestroyMenu(menu);
	if (shellBmp) DeleteObject(shellBmp);
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
