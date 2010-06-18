// Classic Shell (c) 2009-2010, Ivo Beltchev
// The sources for Classic Shell are distributed under the MIT open source license

#include "stdafx.h"
#include "Settings.h"
#include "ExplorerBand.h"
#include "ExplorerBHO.h"
#include "resource.h"
#include "GlobalSettings.h"
#include "TranslationSettings.h"

// Dialog proc for the Settings dialog. Edits the settings and saves them to the registry
INT_PTR CALLBACK SettingsDlgProc( HWND hwndDlg, UINT uMsg, WPARAM wParam, LPARAM lParam )
{
	static HICON s_ShieldIcon;
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
			Sprintf(title,_countof(title),L"Settings for Classic Explorer %d.%d.%d",HIWORD(pVer->dwProductVersionMS),LOWORD(pVer->dwProductVersionMS),HIWORD(pVer->dwProductVersionLS));
		}
		else
			Sprintf(title,_countof(title),L"Settings for Classic Explorer");
		SetWindowText(hwndDlg,title);

		HICON icon=(HICON)LoadImage(g_Instance,MAKEINTRESOURCE(IDI_APPICON),IMAGE_ICON,GetSystemMetrics(SM_CXICON),GetSystemMetrics(SM_CYICON),LR_DEFAULTCOLOR);
		SendMessage(hwndDlg,WM_SETICON,ICON_BIG,(LPARAM)icon);
		icon=(HICON)LoadImage(g_Instance,MAKEINTRESOURCE(IDI_APPICON),IMAGE_ICON,GetSystemMetrics(SM_CXSMICON),GetSystemMetrics(SM_CYSMICON),LR_DEFAULTCOLOR);
		SendMessage(hwndDlg,WM_SETICON,ICON_SMALL,(LPARAM)icon);

		CRegKey regSettings;
		if (regSettings.Open(HKEY_CURRENT_USER,L"Software\\IvoSoft\\ClassicExplorer")!=ERROR_SUCCESS)
			regSettings.Create(HKEY_CURRENT_USER,L"Software\\IvoSoft\\ClassicExplorer");

		DWORD EnableCopyUI, FreeSpace, FoldersSettings, BigButtons, ToolbarButtons, UpButton, AddressBar, SharedOverlay;
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
		if (regSettings.QueryDWORDValue(L"UpButton",UpButton)!=ERROR_SUCCESS)
			UpButton=1;
		if (regSettings.QueryDWORDValue(L"AddressBar",AddressBar)!=ERROR_SUCCESS)
			AddressBar=0;
		if (regSettings.QueryDWORDValue(L"SharedOverlay",SharedOverlay)!=ERROR_SUCCESS)
			SharedOverlay=0;

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

		CheckDlgButton(hwndDlg,IDC_CHECKCOPY,((EnableCopyUI&3)==1 || (EnableCopyUI&3)==2)?BST_CHECKED:BST_UNCHECKED);
		CheckDlgButton(hwndDlg,IDC_CHECKCOPYFOLDER,(EnableCopyUI&1)?BST_CHECKED:BST_UNCHECKED);
		CheckDlgButton(hwndDlg,IDC_CHECKCOPYEXP,(EnableCopyUI&4)?BST_UNCHECKED:BST_CHECKED);
		CheckDlgButton(hwndDlg,IDC_CHECKSIZE,FreeSpace?BST_CHECKED:BST_UNCHECKED);
		CheckDlgButton(hwndDlg,IDC_CHECKBHO,(FoldersSettings&CExplorerBHO::FOLDERS_ALTENTER)?BST_CHECKED:BST_UNCHECKED);
		CheckDlgButton(hwndDlg,IDC_CHECKAUTO,(FoldersSettings&CExplorerBHO::FOLDERS_AUTONAVIGATE)?BST_CHECKED:BST_UNCHECKED);
		CheckDlgButton(hwndDlg,IDC_CHECKNOFADE,(FoldersSettings&CExplorerBHO::FOLDERS_NOFADE)?BST_CHECKED:BST_UNCHECKED);
		CheckDlgButton(hwndDlg,IDC_CHECKOFFSET,(FoldersSettings&CExplorerBHO::FOLDERS_FULLINDENT)?BST_CHECKED:BST_UNCHECKED);

		int style=(FoldersSettings&CExplorerBHO::FOLDERS_STYLE_MASK);
		SendDlgItemMessage(hwndDlg,IDC_COMBOSTYLE,CB_ADDSTRING,0,(LPARAM)L"Windows Vista");
		SendDlgItemMessage(hwndDlg,IDC_COMBOSTYLE,CB_ADDSTRING,0,(LPARAM)L"Windows XP Classic");
		SendDlgItemMessage(hwndDlg,IDC_COMBOSTYLE,CB_ADDSTRING,0,(LPARAM)L"Windows XP Simple");
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

		if (FindSetting("ToolbarItems"))
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
	if (uMsg==WM_INITDIALOG || (uMsg==WM_COMMAND && wParam==IDC_CHECKCOPY))
	{
		BOOL enable=IsDlgButtonChecked(hwndDlg,IDC_CHECKCOPY)==BST_CHECKED;
		EnableWindow(GetDlgItem(hwndDlg,IDC_CHECKCOPYFOLDER),enable);
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

		DWORD EnableCopyUI, FreeSpace, FoldersSettings, BigButtons, ToolbarButtons, UpButton, AddressBar, SharedOverlay;
		if (regSettings.QueryDWORDValue(L"EnableCopyUI",EnableCopyUI)!=ERROR_SUCCESS)
			EnableCopyUI=1;
		if (regSettings.QueryDWORDValue(L"FreeSpace",FreeSpace)!=ERROR_SUCCESS)
			FreeSpace=(LOWORD(GetVersion())==0x0106)?CExplorerBHO::SPACE_SHOW:0;
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
		if (IsDlgButtonChecked(hwndDlg,IDC_CHECKCOPYEXP)!=BST_CHECKED)
			EnableCopyUI2|=4;
		DWORD FreeSpace2=(IsDlgButtonChecked(hwndDlg,IDC_CHECKSIZE)==BST_CHECKED)?CExplorerBHO::SPACE_SHOW:0;
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
		ToolbarButtons2|=((CBandWindow::ID_LAST-1)<<24);
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
			Strcat(path,_countof(path),DOC_PATH L"ClassicExplorer.html");
			ShellExecute(hwndDlg,NULL,path,NULL,NULL,SW_SHOWNORMAL);
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
				TASKDIALOGCONFIG task={sizeof(task),hwndDlg,NULL,TDF_ALLOW_DIALOG_CANCELLATION,TDCBF_OK_BUTTON};
				task.pszMainIcon=TD_INFORMATION_ICON;
				task.pszWindowTitle=L"Classic Explorer";
				task.pszMainInstruction=L"After modifying the ini file you have to open a new Explorer window to use the new settings.\n\nRemember: All lines starting with a semicolon are ignored. Remove the semicolon from the settings you want to use.";
				task.pszVerificationText=L"Don't show this message again";
				BOOL bIgnore=FALSE;
				// call TaskDialogIndirect indirectly because we don't want our DLL to depend on that function. the DLL is loaded even in processes
				// that use the older comctl32.dll, and that dependency causes an error
typedef HRESULT (__stdcall *FTaskDialogIndirect)(const TASKDIALOGCONFIG *pTaskConfig, __out_opt int *pnButton, __out_opt int *pnRadioButton, __out_opt BOOL *pfVerificationFlagChecked);
				FTaskDialogIndirect pTaskDialogIndirect=(FTaskDialogIndirect)GetProcAddress(GetModuleHandle(L"comctl32.dll"),"TaskDialogIndirect");
				if (pTaskDialogIndirect)
				{
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
	INT_PTR res=DialogBox(g_Instance,MAKEINTRESOURCE(IDD_SETTINGS),parent,SettingsDlgProc);
	if (res&2)
		MessageBox(parent,L"You need to log off and back on for the new settings to take effect.",L"Classic Explorer",MB_OK|MB_ICONWARNING);
	else if (res&1)
		MessageBox(parent,L"The new settings will take effect the next time you open an Explorer window.",L"Classic Explorer",MB_OK);
}

void ShowSettingsMenu( HWND parent, int x, int y )
{
	HMENU menu=CreatePopupMenu();
	AppendMenu(menu,MF_STRING,CBandWindow::ID_SETTINGS,FindTranslation("Toolbar.Settings",L"Classic Explorer Settings"));
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
		SetMenuItemInfo(menu,CBandWindow::ID_SETTINGS,FALSE,&mii);
	}
	DeleteDC(hdc);

	DWORD pos=GetMessagePos();
	if (!FindSettingBool("EnableSettings",true))
		EnableMenuItem(menu,0,MF_BYPOSITION|MF_GRAYED);
	int res=TrackPopupMenu(menu,TPM_RETURNCMD,x,y,0,parent,NULL);
	DestroyMenu(menu);
	if (shellBmp) DeleteObject(shellBmp);
	if (res==CBandWindow::ID_SETTINGS)
		ShowSettings(parent);
}
