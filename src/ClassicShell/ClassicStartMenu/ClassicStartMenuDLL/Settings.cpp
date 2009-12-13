// Classic Shell (c) 2009, Ivo Beltchev
// The sources for Classic Shell are distributed under the MIT open source license

#include "stdafx.h"
#include "resource.h"
#include "Settings.h"
#include "ClassicStartMenuDLL.h"

// Default settings
const DWORD DEFAULT_SHOW_FAVORITES=0;
const DWORD DEFAULT_SHOW_DOCUMENTS=1;
const DWORD DEFAULT_SHOW_LOGOFF=0;
const DWORD DEFAULT_SHOW_UNDOCK=1;
const DWORD DEFAULT_EXPAND_CONTROLPANEL=0;
const DWORD DEFAULT_EXPAND_NETWORK=0;
const DWORD DEFAULT_EXPAND_PRINTERS=0;
const DWORD DEFAULT_EXPAND_SHORTCUTS=1;
const DWORD DEFAULT_SMALL_ICONS=0;
const DWORD DEFAULT_THEME=1;
const DWORD DEFAULT_SCROLL_MENUS=0;
const DWORD DEFAULT_CONFIRM_LOGOFF=0;
const DWORD DEFAULT_RECENT_DOCUMENTS=15;
const DWORD DEFAULT_HOTKEY=0;

static HWND g_SettingsDlg;

// Read the settings from the registry
void ReadSettings( StartMenuSettings &settings )
{
	CRegKey regSettings;
	if (regSettings.Open(HKEY_CURRENT_USER,L"Software\\IvoSoft\\ClassicStartMenu")!=ERROR_SUCCESS)
		regSettings.Create(HKEY_CURRENT_USER,L"Software\\IvoSoft\\ClassicStartMenu");

	if (regSettings.QueryDWORDValue(L"ShowFavorites",settings.ShowFavorites)!=ERROR_SUCCESS)
		settings.ShowFavorites=DEFAULT_SHOW_FAVORITES;
	if (regSettings.QueryDWORDValue(L"ShowDocuments",settings.ShowDocuments)!=ERROR_SUCCESS)
		settings.ShowDocuments=DEFAULT_SHOW_DOCUMENTS;
	if (regSettings.QueryDWORDValue(L"ShowLogOff",settings.ShowLogOff)!=ERROR_SUCCESS)
		settings.ShowLogOff=DEFAULT_SHOW_LOGOFF;
	if (regSettings.QueryDWORDValue(L"ShowUndock",settings.ShowUndock)!=ERROR_SUCCESS)
		settings.ShowUndock=DEFAULT_SHOW_UNDOCK;
	if (regSettings.QueryDWORDValue(L"ExpandControlPanel",settings.ExpandControlPanel)!=ERROR_SUCCESS)
		settings.ExpandControlPanel=DEFAULT_EXPAND_CONTROLPANEL;
	if (regSettings.QueryDWORDValue(L"ExpandNetwork",settings.ExpandNetwork)!=ERROR_SUCCESS)
		settings.ExpandNetwork=DEFAULT_EXPAND_NETWORK;
	if (regSettings.QueryDWORDValue(L"ExpandPrinters",settings.ExpandPrinters)!=ERROR_SUCCESS)
		settings.ExpandPrinters=DEFAULT_EXPAND_PRINTERS;
	if (regSettings.QueryDWORDValue(L"ExpandLinks",settings.ExpandLinks)!=ERROR_SUCCESS)
		settings.ExpandLinks=DEFAULT_EXPAND_SHORTCUTS;
	if (regSettings.QueryDWORDValue(L"UseSmallIcons",settings.UseSmallIcons)!=ERROR_SUCCESS)
		settings.UseSmallIcons=DEFAULT_SMALL_ICONS;
	if (regSettings.QueryDWORDValue(L"UseTheme",settings.UseTheme)!=ERROR_SUCCESS)
		settings.UseTheme=DEFAULT_THEME;
	if (regSettings.QueryDWORDValue(L"ScrollMenus",settings.ScrollMenus)!=ERROR_SUCCESS)
		settings.ScrollMenus=DEFAULT_SCROLL_MENUS;
	if (regSettings.QueryDWORDValue(L"ConfirmLogOff",settings.ConfirmLogOff)!=ERROR_SUCCESS)
		settings.ConfirmLogOff=DEFAULT_CONFIRM_LOGOFF;
	if (regSettings.QueryDWORDValue(L"RecentDocuments",settings.RecentDocuments)!=ERROR_SUCCESS)
		settings.RecentDocuments=DEFAULT_RECENT_DOCUMENTS;
	if (regSettings.QueryDWORDValue(L"Hotkey",settings.Hotkey)!=ERROR_SUCCESS)
		settings.Hotkey=DEFAULT_HOTKEY;
}

// Dialog proc for the settings dialog box. Edits and saves the settings.
static INT_PTR CALLBACK SettingsDlgProc( HWND hwndDlg, UINT uMsg, WPARAM wParam, LPARAM lParam )
{
	static DWORD s_Hotkey;
	if (uMsg==WM_INITDIALOG)
	{
		HICON icon=(HICON)LoadImage(g_Instance,MAKEINTRESOURCE(IDI_APPICON),IMAGE_ICON,GetSystemMetrics(SM_CXICON),GetSystemMetrics(SM_CYICON),LR_DEFAULTCOLOR);
		SendMessage(hwndDlg,WM_SETICON,ICON_BIG,(LPARAM)icon);
		icon=(HICON)LoadImage(g_Instance,MAKEINTRESOURCE(IDI_APPICON),IMAGE_ICON,GetSystemMetrics(SM_CXSMICON),GetSystemMetrics(SM_CYSMICON),LR_DEFAULTCOLOR);
		SendMessage(hwndDlg,WM_SETICON,ICON_SMALL,(LPARAM)icon);

		g_SettingsDlg=hwndDlg;

		StartMenuSettings settings;
		ReadSettings(settings);
		s_Hotkey=settings.Hotkey;
		SetHotkey(s_Hotkey==0?0:1);

		CheckDlgButton(hwndDlg,IDC_CHECKFAVORITES,settings.ShowFavorites?BST_CHECKED:BST_UNCHECKED);
		CheckDlgButton(hwndDlg,IDC_CHECKDOCUMENTS,settings.ShowDocuments?BST_CHECKED:BST_UNCHECKED);
		CheckDlgButton(hwndDlg,IDC_CHECKLOGOFF,settings.ShowLogOff?BST_CHECKED:BST_UNCHECKED);
		CheckDlgButton(hwndDlg,IDC_CHECKUNDOCK,settings.ShowUndock?BST_CHECKED:BST_UNCHECKED);
		CheckDlgButton(hwndDlg,IDC_CHECKCONTROLPANEL,settings.ExpandControlPanel?BST_CHECKED:BST_UNCHECKED);
		CheckDlgButton(hwndDlg,IDC_CHECKNETWORK,settings.ExpandNetwork?BST_CHECKED:BST_UNCHECKED);
		CheckDlgButton(hwndDlg,IDC_CHECKPRINTERS,settings.ExpandPrinters?BST_CHECKED:BST_UNCHECKED);
		CheckDlgButton(hwndDlg,IDC_CHECKLINKS,settings.ExpandLinks?BST_CHECKED:BST_UNCHECKED);
		CheckDlgButton(hwndDlg,IDC_CHECKSMALL,settings.UseSmallIcons?BST_CHECKED:BST_UNCHECKED);
		CheckDlgButton(hwndDlg,IDC_CHECKTHEME,settings.UseTheme?BST_CHECKED:BST_UNCHECKED);
		CheckDlgButton(hwndDlg,IDC_CHECKSCROLL,settings.ScrollMenus?BST_CHECKED:BST_UNCHECKED);
		CheckDlgButton(hwndDlg,IDC_CHECKCONFIRM,settings.ConfirmLogOff?BST_CHECKED:BST_UNCHECKED);
		SetDlgItemInt(hwndDlg,IDC_EDITRECENT,settings.RecentDocuments,TRUE);
		CheckDlgButton(hwndDlg,IDC_CHECKHOTKEY,settings.Hotkey!=0?BST_CHECKED:BST_UNCHECKED);
		if (settings.Hotkey!=0 && settings.Hotkey!=1)
			SendDlgItemMessage(hwndDlg,IDC_HOTKEY,HKM_SETHOTKEY,settings.Hotkey,0);
		SendDlgItemMessage(hwndDlg,IDC_HOTKEY,HKM_SETRULES,HKCOMB_NONE|HKCOMB_S|HKCOMB_C,HOTKEYF_CONTROL|HOTKEYF_ALT);
		EnableWindow(GetDlgItem(hwndDlg,IDC_EDITRECENT),settings.ShowDocuments!=0);
		EnableWindow(GetDlgItem(hwndDlg,IDC_CHECKCONFIRM),settings.ShowLogOff!=0);
		EnableWindow(GetDlgItem(hwndDlg,IDC_HOTKEY),settings.Hotkey!=0);

		EnableWindow(GetDlgItem(hwndDlg,IDC_CHECKFAVORITES),!SHRestricted(REST_NOFAVORITESMENU));
		EnableWindow(GetDlgItem(hwndDlg,IDC_CHECKUNDOCK),!SHRestricted(REST_NOSMEJECTPC));

		BOOL bDocuments=!SHRestricted(REST_NORECENTDOCSMENU);
		EnableWindow(GetDlgItem(hwndDlg,IDC_CHECKDOCUMENTS),bDocuments);
		EnableWindow(GetDlgItem(hwndDlg,IDC_EDITRECENT),bDocuments);

		DWORD logoff1=SHRestricted(REST_STARTMENULOGOFF);
		DWORD logoff2=SHRestricted(REST_FORCESTARTMENULOGOFF);
		BOOL bLogOff=!SHRestricted(REST_STARTMENULOGOFF) && !SHRestricted(REST_FORCESTARTMENULOGOFF);
		EnableWindow(GetDlgItem(hwndDlg,IDC_CHECKLOGOFF),logoff1==0 && !logoff2);
		EnableWindow(GetDlgItem(hwndDlg,IDC_CHECKCONFIRM),logoff1!=1);

		bool bNoSetFolders=SHRestricted(REST_NOSETFOLDERS)!=0; // hide control panel, printers, network
		EnableWindow(GetDlgItem(hwndDlg,IDC_CHECKCONTROLPANEL),!bNoSetFolders && !SHRestricted(REST_NOCONTROLPANEL));
		EnableWindow(GetDlgItem(hwndDlg,IDC_CHECKNETWORK),!bNoSetFolders && !SHRestricted(REST_NONETWORKCONNECTIONS));
		EnableWindow(GetDlgItem(hwndDlg,IDC_CHECKPRINTERS),!bNoSetFolders);

		return TRUE;
	}
	if (uMsg==WM_COMMAND && wParam==IDC_CHECKDOCUMENTS)
	{
		EnableWindow(GetDlgItem(hwndDlg,IDC_EDITRECENT),IsDlgButtonChecked(hwndDlg,IDC_CHECKDOCUMENTS)==BST_CHECKED);
		return TRUE;
	}
	if (uMsg==WM_COMMAND && wParam==IDC_CHECKLOGOFF)
	{
		EnableWindow(GetDlgItem(hwndDlg,IDC_CHECKCONFIRM),IsDlgButtonChecked(hwndDlg,IDC_CHECKLOGOFF)==BST_CHECKED);
		return TRUE;
	}
	if (uMsg==WM_COMMAND && wParam==IDC_CHECKHOTKEY)
	{
		EnableWindow(GetDlgItem(hwndDlg,IDC_HOTKEY),IsDlgButtonChecked(hwndDlg,IDC_CHECKHOTKEY)==BST_CHECKED);
		return TRUE;
	}

	if (uMsg==WM_COMMAND && wParam==IDOK)
	{
		CRegKey regSettings;
		if (regSettings.Open(HKEY_CURRENT_USER,L"Software\\IvoSoft\\ClassicStartMenu")!=ERROR_SUCCESS)
			regSettings.Create(HKEY_CURRENT_USER,L"Software\\IvoSoft\\ClassicStartMenu");

		DWORD ShowFavorites=(IsDlgButtonChecked(hwndDlg,IDC_CHECKFAVORITES)==BST_CHECKED)?1:0;
		DWORD ShowDocuments=(IsDlgButtonChecked(hwndDlg,IDC_CHECKDOCUMENTS)==BST_CHECKED)?1:0;
		DWORD ShowLogOff=(IsDlgButtonChecked(hwndDlg,IDC_CHECKLOGOFF)==BST_CHECKED)?1:0;
		DWORD ShowUndock=(IsDlgButtonChecked(hwndDlg,IDC_CHECKUNDOCK)==BST_CHECKED)?1:0;
		DWORD ExpandControlPanel=(IsDlgButtonChecked(hwndDlg,IDC_CHECKCONTROLPANEL)==BST_CHECKED)?1:0;
		DWORD ExpandNetwork=(IsDlgButtonChecked(hwndDlg,IDC_CHECKNETWORK)==BST_CHECKED)?1:0;
		DWORD ExpandPrinters=(IsDlgButtonChecked(hwndDlg,IDC_CHECKPRINTERS)==BST_CHECKED)?1:0;
		DWORD ExpandLinks=(IsDlgButtonChecked(hwndDlg,IDC_CHECKLINKS)==BST_CHECKED)?1:0;
		DWORD UseSmallIcons=(IsDlgButtonChecked(hwndDlg,IDC_CHECKSMALL)==BST_CHECKED)?1:0;
		DWORD UseTheme=(IsDlgButtonChecked(hwndDlg,IDC_CHECKTHEME)==BST_CHECKED)?1:0;
		DWORD ScrollMenus=(IsDlgButtonChecked(hwndDlg,IDC_CHECKSCROLL)==BST_CHECKED)?1:0;
		DWORD ConfirmLogOff=(IsDlgButtonChecked(hwndDlg,IDC_CHECKCONFIRM)==BST_CHECKED)?1:0;
		DWORD RecentDocuments=GetDlgItemInt(hwndDlg,IDC_EDITRECENT,NULL,TRUE);
		DWORD Hotkey=(DWORD)SendDlgItemMessage(hwndDlg,IDC_HOTKEY,HKM_GETHOTKEY,0,0);
		if (IsDlgButtonChecked(hwndDlg,IDC_CHECKHOTKEY)!=BST_CHECKED)
			Hotkey=0;
		else if (Hotkey==0)
			Hotkey=1;

		regSettings.SetDWORDValue(L"ShowFavorites",ShowFavorites);
		regSettings.SetDWORDValue(L"ShowDocuments",ShowDocuments);
		regSettings.SetDWORDValue(L"ShowLogOff",ShowLogOff);
		regSettings.SetDWORDValue(L"ShowUndock",ShowUndock);
		regSettings.SetDWORDValue(L"ExpandControlPanel",ExpandControlPanel);
		regSettings.SetDWORDValue(L"ExpandNetwork",ExpandNetwork);
		regSettings.SetDWORDValue(L"ExpandPrinters",ExpandPrinters);
		regSettings.SetDWORDValue(L"ExpandLinks",ExpandLinks);
		regSettings.SetDWORDValue(L"UseSmallIcons",UseSmallIcons);
		regSettings.SetDWORDValue(L"UseTheme",UseTheme);
		regSettings.SetDWORDValue(L"ScrollMenus",ScrollMenus);
		regSettings.SetDWORDValue(L"ConfirmLogOff",ConfirmLogOff);
		regSettings.SetDWORDValue(L"RecentDocuments",RecentDocuments);
		regSettings.SetDWORDValue(L"Hotkey",Hotkey);
		s_Hotkey=Hotkey;

		DestroyWindow(hwndDlg);
		return TRUE;
	}
	if (uMsg==WM_COMMAND && wParam==IDCANCEL)
	{
		DestroyWindow(hwndDlg);
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
			wcscat_s(path,L"ClassicStartMenu.html");
			ShellExecute(NULL,NULL,path,NULL,NULL,SW_SHOWNORMAL);
			return TRUE;
		}
	}
	if (uMsg==WM_DESTROY)
	{
		SetHotkey(s_Hotkey);
		g_SettingsDlg=NULL;
	}
	return FALSE;
}

// Shows the UI for editing settings
void EditSettings( void )
{
	if (g_SettingsDlg)
		SetActiveWindow(g_SettingsDlg);
	else
	{
		HWND dlg=CreateDialog(g_Instance,MAKEINTRESOURCE(IDD_SETTINGS),NULL,SettingsDlgProc);
		ShowWindow(dlg,SW_SHOWNORMAL);
	}
}

// Close the settings box
void CloseSettings( void )
{
	if (g_SettingsDlg)
		DestroyWindow(g_SettingsDlg);
}

// Process the dialog messages for the settings box
bool IsSettingsMessage( MSG *msg )
{
	if (!g_SettingsDlg) return false;
	if (msg->hwnd!=g_SettingsDlg && !IsChild(g_SettingsDlg,msg->hwnd)) return false;
	// only process keyboard messages. if we process all messages the settings box gets stuck. I don't know why.
	if (msg->message<WM_KEYFIRST || msg->message>WM_KEYLAST) return false;
	return IsDialogMessage(g_SettingsDlg,msg)!=0;
}
