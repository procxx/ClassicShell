// Classic Shell (c) 2009-2010, Ivo Beltchev
// The sources for Classic Shell are distributed under the MIT open source license

#include "stdafx.h"
#include "Settings.h"
#include "ExplorerBand.h"
#include "ExplorerBHO.h"
#include "resource.h"
#include "GlobalSettings.h"
#include "TranslationSettings.h"
#include "dllmain.h"
#include <htmlhelp.h>

// Dialog proc for the Settings dialog. Edits the settings and saves them to the registry
INT_PTR CALLBACK SettingsDlgProc( HWND hwndDlg, UINT uMsg, WPARAM wParam, LPARAM lParam )
{
	static HICON s_ShieldIcon;
	if (uMsg==WM_INITDIALOG)
	{
		wchar_t title[100];
		DWORD ver=GetVersionEx(g_Instance);
		if (ver)
			Sprintf(title,_countof(title),LoadStringEx(IDS_SETTINGS_TITLE_VER),ver>>24,(ver>>16)&0xFF,ver&0xFFFF);
		else
			Sprintf(title,_countof(title),LoadStringEx(IDS_SETTINGS_TITLE));
		SetWindowText(hwndDlg,title);

		HICON icon=(HICON)LoadImage(g_Instance,MAKEINTRESOURCE(IDI_APPICON),IMAGE_ICON,GetSystemMetrics(SM_CXICON),GetSystemMetrics(SM_CYICON),LR_DEFAULTCOLOR);
		SendMessage(hwndDlg,WM_SETICON,ICON_BIG,(LPARAM)icon);
		icon=(HICON)LoadImage(g_Instance,MAKEINTRESOURCE(IDI_APPICON),IMAGE_ICON,GetSystemMetrics(SM_CXSMICON),GetSystemMetrics(SM_CYSMICON),LR_DEFAULTCOLOR);
		SendMessage(hwndDlg,WM_SETICON,ICON_SMALL,(LPARAM)icon);

		CRegKey regSettings;
		if (regSettings.Open(HKEY_CURRENT_USER,L"Software\\IvoSoft\\ClassicExplorer")!=ERROR_SUCCESS)
			regSettings.Create(HKEY_CURRENT_USER,L"Software\\IvoSoft\\ClassicExplorer");

		DWORD EnableFileUI, FreeSpace, FoldersSettings, BigButtons, ToolbarButtons, UpButton, AddressBar, SharedOverlay;
		if (regSettings.QueryDWORDValue(L"EnableFileUI",EnableFileUI)!=ERROR_SUCCESS)
			EnableFileUI=FILEUI_DEFAULT;
		if (regSettings.QueryDWORDValue(L"FreeSpace",FreeSpace)!=ERROR_SUCCESS)
			FreeSpace=(LOWORD(GetVersion())==0x0006)?0:CExplorerBHO::SPACE_SHOW;
		if (regSettings.QueryDWORDValue(L"FoldersSettings",FoldersSettings)!=ERROR_SUCCESS)
			FoldersSettings=CExplorerBHO::FOLDERS_DEFAULT;
		if (regSettings.QueryDWORDValue(L"BigButtons",BigButtons)!=ERROR_SUCCESS)
			BigButtons=0;
		if (regSettings.QueryDWORDValue(L"ToolbarButtons",ToolbarButtons)!=ERROR_SUCCESS)
			ToolbarButtons=DEFAULT_BUTTONS|((CBandWindow::ID_LAST_STD-1)<<24);
		if (regSettings.QueryDWORDValue(L"UpButton",UpButton)!=ERROR_SUCCESS)
			UpButton=1;
		if (regSettings.QueryDWORDValue(L"AddressBar",AddressBar)!=ERROR_SUCCESS)
			AddressBar=0;
		if (regSettings.QueryDWORDValue(L"SharedOverlay",SharedOverlay)!=ERROR_SUCCESS)
			SharedOverlay=0;

		if (!(ToolbarButtons&0xFF000000)) ToolbarButtons|=0x07000002; // for backwards compatibility (when there were 7 buttons the the button count was not saved)
		unsigned int mask1=(((2<<(ToolbarButtons>>24))-1)&~1); // bits to keep
		unsigned int mask2=(((2<<CBandWindow::ID_LAST_STD)-1)&~1)&~mask1; // bits to replace with defaults
		ToolbarButtons=(ToolbarButtons&mask1)|(DEFAULT_BUTTONS&mask2)|((CBandWindow::ID_LAST_STD-1)<<24);
		if ((ToolbarButtons&0xFFFFFF)==0)
			ToolbarButtons|=1<<DEFAULT_ONLY_BUTTON;

		HWND parent=GetParent(hwndDlg);
		if (parent)
		{
			RECT rc1,rc2;
			GetWindowRect(hwndDlg,&rc1);
			GetWindowRect(parent,&rc2);
			OffsetRect(&rc1,(rc2.left+rc2.right)/2-(rc1.left+rc1.right)/2,(rc2.top+rc2.bottom)/2-(rc1.top+rc1.bottom)/2);
			if (rc1.top<rc2.top) OffsetRect(&rc1,0,rc2.top-rc1.top);
			SetWindowPos(hwndDlg,NULL,rc1.left,rc1.top,rc1.right-rc1.left,rc1.bottom-rc1.top,SWP_NOZORDER);
			SendMessage(hwndDlg,DM_REPOSITION,0,0);
		}

		CheckDlgButton(hwndDlg,IDC_CHECKCOPY,(EnableFileUI&FILEUI_FILE)?BST_CHECKED:BST_UNCHECKED);
		CheckDlgButton(hwndDlg,IDC_CHECKCOPYFOLDER,(EnableFileUI&FILEUI_FOLDER)?BST_CHECKED:BST_UNCHECKED);
		CheckDlgButton(hwndDlg,IDC_CHECKMORE,(EnableFileUI&FILEUI_MORE)?BST_CHECKED:BST_UNCHECKED);
		CheckDlgButton(hwndDlg,IDC_CHECKCOPYEXP,(EnableFileUI&FILEUI_OTHERAPPS)?BST_UNCHECKED:BST_CHECKED);
		CheckDlgButton(hwndDlg,IDC_CHECKSIZE,(FreeSpace&CExplorerBHO::SPACE_SHOW)?BST_CHECKED:BST_UNCHECKED);
		CheckDlgButton(hwndDlg,IDC_CHECKINFO,(FreeSpace&CExplorerBHO::SPACE_INFOTIP)?BST_UNCHECKED:BST_CHECKED);
		CheckDlgButton(hwndDlg,IDC_CHECKBHO,(FoldersSettings&CExplorerBHO::FOLDERS_ALTENTER)?BST_CHECKED:BST_UNCHECKED);
		CheckDlgButton(hwndDlg,IDC_CHECKAUTO,(FoldersSettings&CExplorerBHO::FOLDERS_AUTONAVIGATE)?BST_CHECKED:BST_UNCHECKED);
		CheckDlgButton(hwndDlg,IDC_CHECKNOFADE,(FoldersSettings&CExplorerBHO::FOLDERS_NOFADE)?BST_CHECKED:BST_UNCHECKED);
		CheckDlgButton(hwndDlg,IDC_CHECKOFFSET,(FoldersSettings&CExplorerBHO::FOLDERS_FULLINDENT)?BST_CHECKED:BST_UNCHECKED);

		int style=(FoldersSettings&CExplorerBHO::FOLDERS_STYLE_MASK);
		SendDlgItemMessage(hwndDlg,IDC_COMBOSTYLE,CB_ADDSTRING,0,(LPARAM)(const wchar_t*)LoadStringEx(IDS_STYLE_VISTA));
		SendDlgItemMessage(hwndDlg,IDC_COMBOSTYLE,CB_ADDSTRING,0,(LPARAM)(const wchar_t*)LoadStringEx(IDS_STYLE_XP_CLASSIC));
		SendDlgItemMessage(hwndDlg,IDC_COMBOSTYLE,CB_ADDSTRING,0,(LPARAM)(const wchar_t*)LoadStringEx(IDS_STYLE_XP_SIMPLE));
		if (style==CExplorerBHO::FOLDERS_CLASSIC)
			SendDlgItemMessage(hwndDlg,IDC_COMBOSTYLE,CB_SETCURSEL,1,0);
		else if (style==CExplorerBHO::FOLDERS_SIMPLE)
			SendDlgItemMessage(hwndDlg,IDC_COMBOSTYLE,CB_SETCURSEL,2,0);
		else
			SendDlgItemMessage(hwndDlg,IDC_COMBOSTYLE,CB_SETCURSEL,0,0);

		CheckDlgButton(hwndDlg,IDC_CHECKUP,UpButton?BST_CHECKED:BST_UNCHECKED);
		CheckDlgButton(hwndDlg,IDC_CHECKTITLE,(AddressBar&CExplorerBHO::ADDRESS_SHOWTITLE)?BST_CHECKED:BST_UNCHECKED);
		CheckDlgButton(hwndDlg,IDC_CHECKICON,(AddressBar&CExplorerBHO::ADDRESS_SHOWICON)?BST_CHECKED:BST_UNCHECKED);
		CheckDlgButton(hwndDlg,IDC_CHECKCRUMBS,(AddressBar&CExplorerBHO::ADDRESS_NOBREADCRUMBS)?BST_CHECKED:BST_UNCHECKED);
		CheckDlgButton(hwndDlg,IDC_CHECKSHARED,SharedOverlay?BST_CHECKED:BST_UNCHECKED);
		CheckDlgButton(hwndDlg,IDC_CHECKSHAREDEXP,SharedOverlay!=2?BST_CHECKED:BST_UNCHECKED);

		CheckDlgButton(hwndDlg,IDC_CHECKBIG,BigButtons?BST_CHECKED:BST_UNCHECKED);
		CheckDlgButton(hwndDlg,IDC_CHECK1,(ToolbarButtons&(1<<CBandWindow::ID_GOUP))?BST_CHECKED:BST_UNCHECKED);
		CheckDlgButton(hwndDlg,IDC_CHECK2,(ToolbarButtons&(1<<CBandWindow::ID_CUT))?BST_CHECKED:BST_UNCHECKED);
		CheckDlgButton(hwndDlg,IDC_CHECK3,(ToolbarButtons&(1<<CBandWindow::ID_COPY))?BST_CHECKED:BST_UNCHECKED);
		CheckDlgButton(hwndDlg,IDC_CHECK4,(ToolbarButtons&(1<<CBandWindow::ID_PASTE))?BST_CHECKED:BST_UNCHECKED);
		CheckDlgButton(hwndDlg,IDC_CHECK5,(ToolbarButtons&(1<<CBandWindow::ID_DELETE))?BST_CHECKED:BST_UNCHECKED);
		CheckDlgButton(hwndDlg,IDC_CHECK6,(ToolbarButtons&(1<<CBandWindow::ID_PROPERTIES))?BST_CHECKED:BST_UNCHECKED);
		CheckDlgButton(hwndDlg,IDC_CHECK7,(ToolbarButtons&(1<<CBandWindow::ID_EMAIL))?BST_CHECKED:BST_UNCHECKED);
		CheckDlgButton(hwndDlg,IDC_CHECK8,(ToolbarButtons&(1<<CBandWindow::ID_SETTINGS))?BST_CHECKED:BST_UNCHECKED);

		if (FindSettingBool("EnableCustomToolbar",false))
		{
			EnableWindow(GetDlgItem(hwndDlg,IDC_CHECK1),FALSE);
			EnableWindow(GetDlgItem(hwndDlg,IDC_CHECK2),FALSE);
			EnableWindow(GetDlgItem(hwndDlg,IDC_CHECK3),FALSE);
			EnableWindow(GetDlgItem(hwndDlg,IDC_CHECK4),FALSE);
			EnableWindow(GetDlgItem(hwndDlg,IDC_CHECK5),FALSE);
			EnableWindow(GetDlgItem(hwndDlg,IDC_CHECK6),FALSE);
			EnableWindow(GetDlgItem(hwndDlg,IDC_CHECK7),FALSE);
			EnableWindow(GetDlgItem(hwndDlg,IDC_CHECK8),FALSE);
		}

		SHSTOCKICONINFO sii={sizeof(sii)};
		SHGetStockIconInfo(SIID_SHIELD,SHGSI_ICON|SHGSI_SMALLICON,&sii);
		s_ShieldIcon=sii.hIcon;
		HWND shield=GetDlgItem(hwndDlg,IDC_SHIELD);
		RECT rc;
		GetWindowRect(shield,&rc);
		POINT pt={rc.right,rc.bottom};
		ScreenToClient(hwndDlg,&pt);
		int iconSize=GetSystemMetrics(SM_CXSMICON);
		SetWindowPos(shield,NULL,pt.x-iconSize,pt.y-iconSize,iconSize,iconSize,SWP_NOZORDER);
		SendMessage(shield,STM_SETICON,(WPARAM)sii.hIcon,0);
	}
	if (uMsg==WM_INITDIALOG || (uMsg==WM_COMMAND && wParam==MAKEWPARAM(IDC_COMBOSTYLE,CBN_SELENDOK)))
	{
		int style=(int)SendDlgItemMessage(hwndDlg,IDC_COMBOSTYLE,CB_GETCURSEL,0,0);
		EnableWindow(GetDlgItem(hwndDlg,IDC_CHECKNOFADE),style!=1);
		EnableWindow(GetDlgItem(hwndDlg,IDC_CHECKOFFSET),style!=1);
		if (uMsg==WM_COMMAND)
			return TRUE;
	}
	if (uMsg==WM_INITDIALOG || (uMsg==WM_COMMAND && wParam==IDC_CHECKSIZE))
	{
		EnableWindow(GetDlgItem(hwndDlg,IDC_CHECKINFO),IsDlgButtonChecked(hwndDlg,IDC_CHECKSIZE)==BST_CHECKED);
		if (uMsg==WM_COMMAND)
			return TRUE;
	}
	if (uMsg==WM_INITDIALOG || (uMsg==WM_COMMAND && (wParam==IDC_CHECKCOPY || wParam==IDC_CHECKCOPYFOLDER || wParam==IDC_CHECKMORE)))
	{
		BOOL enable=IsDlgButtonChecked(hwndDlg,IDC_CHECKCOPY)==BST_CHECKED ||
				IsDlgButtonChecked(hwndDlg,IDC_CHECKCOPYFOLDER)==BST_CHECKED ||
				IsDlgButtonChecked(hwndDlg,IDC_CHECKMORE)==BST_CHECKED;
		EnableWindow(GetDlgItem(hwndDlg,IDC_CHECKCOPYEXP),enable);
		if (uMsg==WM_COMMAND)
			return TRUE;
	}
	if (uMsg==WM_INITDIALOG || (uMsg==WM_COMMAND && wParam==IDC_CHECKSHARED))
	{
		EnableWindow(GetDlgItem(hwndDlg,IDC_CHECKSHAREDEXP),IsDlgButtonChecked(hwndDlg,IDC_CHECKSHARED)==BST_CHECKED);
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

		DWORD EnableFileUI, FreeSpace, FoldersSettings, BigButtons, ToolbarButtons, UpButton, AddressBar, SharedOverlay;
		if (regSettings.QueryDWORDValue(L"EnableFileUI",EnableFileUI)!=ERROR_SUCCESS)
			EnableFileUI=FILEUI_DEFAULT;
		if (regSettings.QueryDWORDValue(L"FreeSpace",FreeSpace)!=ERROR_SUCCESS)
			FreeSpace=(LOWORD(GetVersion())==0x0006)?0:CExplorerBHO::SPACE_SHOW;
		if (regSettings.QueryDWORDValue(L"FoldersSettings",FoldersSettings)!=ERROR_SUCCESS)
			FoldersSettings=CExplorerBHO::FOLDERS_DEFAULT;
		if (regSettings.QueryDWORDValue(L"BigButtons",BigButtons)!=ERROR_SUCCESS)
			BigButtons=0;
		if (regSettings.QueryDWORDValue(L"UpButton",UpButton)!=ERROR_SUCCESS)
			UpButton=1;
		if (regSettings.QueryDWORDValue(L"AddressBar",AddressBar)!=ERROR_SUCCESS)
			AddressBar=0;
		if (regSettings.QueryDWORDValue(L"SharedOverlay",SharedOverlay)!=ERROR_SUCCESS)
			SharedOverlay=0;
		if (regSettings.QueryDWORDValue(L"ToolbarButtons",ToolbarButtons)!=ERROR_SUCCESS)
			ToolbarButtons=DEFAULT_BUTTONS|((CBandWindow::ID_LAST_STD-1)<<24);
		DWORD ToolbarButtons0=ToolbarButtons;
		if (!(ToolbarButtons&0xFF000000)) ToolbarButtons|=0x07000002; // for backwards compatibility (when there were 7 buttons the the button count was not saved)
		unsigned int mask1=(((2<<(ToolbarButtons>>24))-1)&~1); // bits to keep
		unsigned int mask2=(((2<<CBandWindow::ID_LAST_STD)-1)&~1)&~mask1; // bits to replace with defaults
		ToolbarButtons=(ToolbarButtons&mask1)|(DEFAULT_BUTTONS&mask2)|((CBandWindow::ID_LAST_STD-1)<<24);
		if ((ToolbarButtons&0xFFFFFF)==0)
			ToolbarButtons|=1<<DEFAULT_ONLY_BUTTON;

		DWORD EnableFileUI2=(IsDlgButtonChecked(hwndDlg,IDC_CHECKCOPY)==BST_CHECKED)?FILEUI_FILE :0;
		if (IsDlgButtonChecked(hwndDlg,IDC_CHECKCOPYFOLDER)==BST_CHECKED)
			EnableFileUI2|=FILEUI_FOLDER;
		if (IsDlgButtonChecked(hwndDlg,IDC_CHECKMORE)==BST_CHECKED)
			EnableFileUI2|=FILEUI_MORE;
		if (IsDlgButtonChecked(hwndDlg,IDC_CHECKCOPYEXP)!=BST_CHECKED)
			EnableFileUI2|=FILEUI_OTHERAPPS;
		DWORD FreeSpace2=(IsDlgButtonChecked(hwndDlg,IDC_CHECKSIZE)==BST_CHECKED)?CExplorerBHO::SPACE_SHOW:0;
		if (IsDlgButtonChecked(hwndDlg,IDC_CHECKINFO)!=BST_CHECKED)
			FreeSpace2|=CExplorerBHO::SPACE_INFOTIP;
		DWORD FoldersSettings2=0;
		if (IsDlgButtonChecked(hwndDlg,IDC_CHECKBHO)==BST_CHECKED)
			FoldersSettings2|=CExplorerBHO::FOLDERS_ALTENTER;
		if (IsDlgButtonChecked(hwndDlg,IDC_CHECKAUTO)==BST_CHECKED)
			FoldersSettings2|=CExplorerBHO::FOLDERS_AUTONAVIGATE;
		int style=(int)SendDlgItemMessage(hwndDlg,IDC_COMBOSTYLE,CB_GETCURSEL,0,0);
		if (style==0) FoldersSettings2|=CExplorerBHO::FOLDERS_VISTA;
		if (style==1) FoldersSettings2|=CExplorerBHO::FOLDERS_CLASSIC;
		if (style==2) FoldersSettings2|=CExplorerBHO::FOLDERS_SIMPLE;
		if (IsDlgButtonChecked(hwndDlg,IDC_CHECKNOFADE)==BST_CHECKED)
			FoldersSettings2|=CExplorerBHO::FOLDERS_NOFADE;
		if (IsDlgButtonChecked(hwndDlg,IDC_CHECKOFFSET)==BST_CHECKED)
			FoldersSettings2|=CExplorerBHO::FOLDERS_FULLINDENT;

		DWORD UpButton2=(IsDlgButtonChecked(hwndDlg,IDC_CHECKUP)==BST_CHECKED)?1:0;
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
		ToolbarButtons2|=((CBandWindow::ID_LAST_STD-1)<<24);
		if ((ToolbarButtons2&0xFFFFFF)==0)
			ToolbarButtons2|=1<<DEFAULT_ONLY_BUTTON;

		DWORD AddressBar2=0;
		if (IsDlgButtonChecked(hwndDlg,IDC_CHECKTITLE)==BST_CHECKED)
			AddressBar2|=CExplorerBHO::ADDRESS_SHOWTITLE;
		if (IsDlgButtonChecked(hwndDlg,IDC_CHECKICON)==BST_CHECKED)
			AddressBar2|=CExplorerBHO::ADDRESS_SHOWICON;
		if (IsDlgButtonChecked(hwndDlg,IDC_CHECKCRUMBS)==BST_CHECKED)
			AddressBar2|=CExplorerBHO::ADDRESS_NOBREADCRUMBS;

		DWORD SharedOverlay2=0;
		if (IsDlgButtonChecked(hwndDlg,IDC_CHECKSHARED)==BST_CHECKED)
			SharedOverlay2=IsDlgButtonChecked(hwndDlg,IDC_CHECKSHAREDEXP)==BST_CHECKED?1:2;

		int res=0;
		if (EnableFileUI!=EnableFileUI2)
		{
extern bool g_bHookCopyThreads;
			if (!g_bHookCopyThreads && (EnableFileUI&FILEUI_ALL)!=(EnableFileUI2&FILEUI_ALL))
				res|=2;
			regSettings.SetDWORDValue(L"EnableFileUI",EnableFileUI2);
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
		if (UpButton!=UpButton2)
		{
			regSettings.SetDWORDValue(L"UpButton",UpButton2);
			res|=1;
		}
		if (AddressBar!=AddressBar2)
		{
			regSettings.SetDWORDValue(L"AddressBar",AddressBar2);
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

		if (SharedOverlay!=SharedOverlay2)
		{
			regSettings.SetDWORDValue(L"SharedOverlay",SharedOverlay2);
			res|=2;
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
			Strcat(path,_countof(path),DOC_PATH L"ClassicShell.chm::/ClassicExplorer.html");
			HtmlHelp(GetDesktopWindow(),path,HH_DISPLAY_TOPIC,NULL);
			return TRUE;
		}
		if (pHdr->idFrom==IDC_LINKINI && (pHdr->code==NM_CLICK || pHdr->code==NM_RETURN))
		{
			CRegKey regSettings;
			if (regSettings.Open(HKEY_CURRENT_USER,L"Software\\IvoSoft\\ClassicExplorer")!=ERROR_SUCCESS)
				regSettings.Create(HKEY_CURRENT_USER,L"Software\\IvoSoft\\ClassicExplorer");

			DWORD show;
			if (regSettings.QueryDWORDValue(L"IgnoreIniWarning",show)!=ERROR_SUCCESS)
			{
				// call TaskDialogIndirect indirectly because we don't want our DLL to depend on that function. the DLL is loaded even in processes
				// that use the older comctl32.dll, and that dependency causes an error
typedef HRESULT (__stdcall *FTaskDialogIndirect)(const TASKDIALOGCONFIG *pTaskConfig, __out_opt int *pnButton, __out_opt int *pnRadioButton, __out_opt BOOL *pfVerificationFlagChecked);
				FTaskDialogIndirect pTaskDialogIndirect=(FTaskDialogIndirect)GetProcAddress(GetModuleHandle(L"comctl32.dll"),"TaskDialogIndirect");
				if (pTaskDialogIndirect)
				{
					TASKDIALOGCONFIG task={sizeof(task),hwndDlg,NULL,TDF_ALLOW_DIALOG_CANCELLATION,TDCBF_OK_BUTTON};
					CString title=LoadStringEx(IDS_APP_TITLE);
					CString warning=LoadStringEx(IDS_INI_WARNING);
					CString ignore=LoadStringEx(IDS_INI_IGNORE);
					task.pszMainIcon=TD_INFORMATION_ICON;
					task.pszWindowTitle=title;
					task.pszMainInstruction=warning;
					task.pszVerificationText=ignore;
					BOOL bIgnore=FALSE;

					pTaskDialogIndirect(&task,NULL,NULL,&bIgnore);
					if (bIgnore)
						regSettings.SetDWORDValue(L"IgnoreIniWarning",1);
				}
			}

			// run Notepad as administrator (the ini file is most likely in a protected folder)
			wchar_t path[_MAX_PATH];
			GetModuleFileName(g_Instance,path,_countof(path));
			*PathFindFileName(path)=0;
			Strcat(path,_countof(path),INI_PATH L"Explorer.ini");
			ShellExecute(hwndDlg,L"runas",L"notepad",path,NULL,SW_SHOWNORMAL);
			return TRUE;
		}
	}
	if (uMsg==WM_DESTROY)
	{
		if (s_ShieldIcon) DestroyIcon(s_ShieldIcon);
		s_ShieldIcon=NULL;
	}
	return FALSE;
}

void ShowSettings( HWND parent )
{
	INT_PTR res=RunSettingsDialog(parent,SettingsDlgProc);
	if (res&2)
		MessageBox(parent,LoadStringEx(IDS_NEW_SETTINGS2),LoadStringEx(IDS_APP_TITLE),MB_OK|MB_ICONWARNING);
	else if (res&1)
		MessageBox(parent,LoadStringEx(IDS_NEW_SETTINGS1),LoadStringEx(IDS_APP_TITLE),MB_OK);
}

void ShowSettingsMenu( HWND parent, int x, int y )
{
	HMENU menu=CreatePopupMenu();
	AppendMenu(menu,MF_STRING,CBandWindow::ID_SETTINGS,FindTranslation("Toolbar.Settings",L"Classic Explorer Settings"));
	int size=16;
	HBITMAP shellBmp=NULL;
	HICON icon=(HICON)LoadImage(g_Instance,MAKEINTRESOURCE(IDI_APPICON),IMAGE_ICON,size,size,LR_DEFAULTCOLOR);
	if (icon)
	{
		shellBmp=BitmapFromIcon(icon,size,NULL,true);
		MENUITEMINFO mii={sizeof(mii)};
		mii.fMask=MIIM_BITMAP;
		mii.hbmpItem=shellBmp;
		SetMenuItemInfo(menu,CBandWindow::ID_SETTINGS,FALSE,&mii);
	}
	MENUINFO info={sizeof(info),MIM_STYLE,MNS_CHECKORBMP};
	SetMenuInfo(menu,&info);

	DWORD pos=GetMessagePos();
	if (!FindSettingBool("EnableSettings",true))
		EnableMenuItem(menu,0,MF_BYPOSITION|MF_GRAYED);
	int res=TrackPopupMenu(menu,TPM_RETURNCMD,x,y,0,parent,NULL);
	DestroyMenu(menu);
	if (shellBmp) DeleteObject(shellBmp);
	if (res==CBandWindow::ID_SETTINGS)
		ShowSettings(parent);
}

void WINAPI ShowExplorerSettings( void )
{
	ShowSettings(NULL);
}
