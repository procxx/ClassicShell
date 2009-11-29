// Classic Shell (c) 2009, Ivo Beltchev
// The sources for Classic Shell are distributed under the MIT open source license

#include "stdafx.h"
#include "resource.h"
#include "Settings.h"

// Default settings
const DWORD DEFAULT_SHOW_FAVORITES=0;
const DWORD DEFAULT_SHOW_LOGOFF=0;
const DWORD DEFAULT_SHOW_UNDOCK=1;
const DWORD DEFAULT_EXPAND_CONTROLPANEL=0;
const DWORD DEFAULT_EXPAND_NETWORK=0;
const DWORD DEFAULT_EXPAND_PRINTERS=0;
const DWORD DEFAULT_SMALL_ICONS=0;
const DWORD DEFAULT_SCROLL_MENUS=0;
const DWORD DEFAULT_RECENT_DOCUMENTS=15;

static HWND g_SettingsDlg;

// Read the settings from the registry
void ReadSettings( StartMenuSettings &settings )
{
	CRegKey regSettings;
	if (regSettings.Open(HKEY_CURRENT_USER,L"Software\\IvoSoft\\ClassicStartMenu")!=ERROR_SUCCESS)
		regSettings.Create(HKEY_CURRENT_USER,L"Software\\IvoSoft\\ClassicStartMenu");

	if (regSettings.QueryDWORDValue(L"ShowFavorites",settings.ShowFavorites)!=ERROR_SUCCESS)
		settings.ShowFavorites=DEFAULT_SHOW_FAVORITES;
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
	if (regSettings.QueryDWORDValue(L"UseSmallIcons",settings.UseSmallIcons)!=ERROR_SUCCESS)
		settings.UseSmallIcons=DEFAULT_SMALL_ICONS;
	if (regSettings.QueryDWORDValue(L"ScrollMenus",settings.ScrollMenus)!=ERROR_SUCCESS)
		settings.ScrollMenus=DEFAULT_SCROLL_MENUS;
	if (regSettings.QueryDWORDValue(L"RecentDocuments",settings.RecentDocuments)!=ERROR_SUCCESS)
		settings.RecentDocuments=DEFAULT_RECENT_DOCUMENTS;
}

// Dialog proc for the settings dialog box. Edits and saves the settings.
static INT_PTR CALLBACK SettingsDlgProc( HWND hwndDlg, UINT uMsg, WPARAM wParam, LPARAM lParam )
{
	if (uMsg==WM_INITDIALOG)
	{
		g_SettingsDlg=hwndDlg;

		StartMenuSettings settings;
		ReadSettings(settings);

		CheckDlgButton(hwndDlg,IDC_CHECKFAVORITES,settings.ShowFavorites?BST_CHECKED:BST_UNCHECKED);
		CheckDlgButton(hwndDlg,IDC_CHECKLOGOFF,settings.ShowLogOff?BST_CHECKED:BST_UNCHECKED);
		CheckDlgButton(hwndDlg,IDC_CHECKUNDOCK,settings.ShowUndock?BST_CHECKED:BST_UNCHECKED);
		CheckDlgButton(hwndDlg,IDC_CHECKCONTROLPANEL,settings.ExpandControlPanel?BST_CHECKED:BST_UNCHECKED);
		CheckDlgButton(hwndDlg,IDC_CHECKNETWORK,settings.ExpandNetwork?BST_CHECKED:BST_UNCHECKED);
		CheckDlgButton(hwndDlg,IDC_CHECKPRINTERS,settings.ExpandPrinters?BST_CHECKED:BST_UNCHECKED);
		CheckDlgButton(hwndDlg,IDC_CHECKSMALL,settings.UseSmallIcons?BST_CHECKED:BST_UNCHECKED);
		CheckDlgButton(hwndDlg,IDC_CHECKSCROLL,settings.ScrollMenus?BST_CHECKED:BST_UNCHECKED);
		SetDlgItemInt(hwndDlg,IDC_EDITRECENT,settings.RecentDocuments,TRUE);
		return TRUE;
	}
	if (uMsg==WM_COMMAND && wParam==IDOK)
	{
		CRegKey regSettings;
		if (regSettings.Open(HKEY_CURRENT_USER,L"Software\\IvoSoft\\ClassicStartMenu")!=ERROR_SUCCESS)
			regSettings.Create(HKEY_CURRENT_USER,L"Software\\IvoSoft\\ClassicStartMenu");

		DWORD ShowFavorites=(IsDlgButtonChecked(hwndDlg,IDC_CHECKFAVORITES)==BST_CHECKED)?1:0;
		DWORD ShowLogOff=(IsDlgButtonChecked(hwndDlg,IDC_CHECKLOGOFF)==BST_CHECKED)?1:0;
		DWORD ShowUndock=(IsDlgButtonChecked(hwndDlg,IDC_CHECKUNDOCK)==BST_CHECKED)?1:0;
		DWORD ExpandControlPanel=(IsDlgButtonChecked(hwndDlg,IDC_CHECKCONTROLPANEL)==BST_CHECKED)?1:0;
		DWORD ExpandNetwork=(IsDlgButtonChecked(hwndDlg,IDC_CHECKNETWORK)==BST_CHECKED)?1:0;
		DWORD ExpandPrinters=(IsDlgButtonChecked(hwndDlg,IDC_CHECKPRINTERS)==BST_CHECKED)?1:0;
		DWORD UseSmallIcons=(IsDlgButtonChecked(hwndDlg,IDC_CHECKSMALL)==BST_CHECKED)?1:0;
		DWORD ScrollMenus=(IsDlgButtonChecked(hwndDlg,IDC_CHECKSCROLL)==BST_CHECKED)?1:0;
		DWORD RecentDocuments=GetDlgItemInt(hwndDlg,IDC_EDITRECENT,NULL,TRUE);

		regSettings.SetDWORDValue(L"ShowFavorites",ShowFavorites);
		regSettings.SetDWORDValue(L"ShowLogOff",ShowLogOff);
		regSettings.SetDWORDValue(L"ShowUndock",ShowUndock);
		regSettings.SetDWORDValue(L"ExpandControlPanel",ExpandControlPanel);
		regSettings.SetDWORDValue(L"ExpandNetwork",ExpandNetwork);
		regSettings.SetDWORDValue(L"ExpandPrinters",ExpandPrinters);
		regSettings.SetDWORDValue(L"UseSmallIcons",UseSmallIcons);
		regSettings.SetDWORDValue(L"ScrollMenus",ScrollMenus);
		regSettings.SetDWORDValue(L"RecentDocuments",RecentDocuments);

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
		g_SettingsDlg=NULL;
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
