// Classic Shell (c) 2009-2010, Ivo Beltchev
// The sources for Classic Shell are distributed under the MIT open source license

#include "stdafx.h"
#include "resource.h"
#include "Settings.h"
#include "ClassicStartMenuDLL.h"
#include "SkinManager.h"
#include <uxtheme.h>
#include <dwmapi.h>

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
const DWORD DEFAULT_HOTKEY=0;

static HWND g_SettingsDlg;

static void ReadSettings( HWND hwndDlg, StartMenuSettings &settings )
{
	settings.ShowFavorites=(IsDlgButtonChecked(hwndDlg,IDC_CHECKFAVORITES)==BST_CHECKED)?1:0;
	settings.ShowDocuments=(IsDlgButtonChecked(hwndDlg,IDC_CHECKDOCUMENTS)==BST_CHECKED)?1:0;
	settings.ShowLogOff=(IsDlgButtonChecked(hwndDlg,IDC_CHECKLOGOFF)==BST_CHECKED)?1:0;
	settings.ShowUndock=(IsDlgButtonChecked(hwndDlg,IDC_CHECKUNDOCK)==BST_CHECKED)?1:0;
	settings.ExpandControlPanel=(IsDlgButtonChecked(hwndDlg,IDC_CHECKCONTROLPANEL)==BST_CHECKED)?1:0;
	settings.ExpandNetwork=(IsDlgButtonChecked(hwndDlg,IDC_CHECKNETWORK)==BST_CHECKED)?1:0;
	settings.ExpandPrinters=(IsDlgButtonChecked(hwndDlg,IDC_CHECKPRINTERS)==BST_CHECKED)?1:0;
	settings.ExpandLinks=(IsDlgButtonChecked(hwndDlg,IDC_CHECKLINKS)==BST_CHECKED)?1:0;
	settings.ScrollMenus=(IsDlgButtonChecked(hwndDlg,IDC_CHECKSCROLL)==BST_CHECKED)?1:0;
	settings.Hotkey=(DWORD)SendDlgItemMessage(hwndDlg,IDC_HOTKEY,HKM_GETHOTKEY,0,0);
	if (IsDlgButtonChecked(hwndDlg,IDC_CHECKHOTKEY)!=BST_CHECKED)
		settings.Hotkey=0;
	else if (settings.Hotkey==0)
		settings.Hotkey=1;

	wchar_t skinName[256];
	int idx=(int)SendDlgItemMessage(hwndDlg,IDC_COMBOSKIN,CB_GETCURSEL,0,0);
	SendDlgItemMessage(hwndDlg,IDC_COMBOSKIN,CB_GETLBTEXT,idx,(LPARAM)skinName);
	if (wcscmp(skinName,L"<Default>")==0)
		settings.SkinName.Empty();
	else
		settings.SkinName=skinName;

	idx=(int)SendDlgItemMessage(hwndDlg,IDC_COMBOVAR,CB_GETCURSEL,0,0);
	if (idx>=0)
	{
		SendDlgItemMessage(hwndDlg,IDC_COMBOVAR,CB_GETLBTEXT,idx,(LPARAM)skinName);
		settings.SkinVariation=skinName;
	}
	else
		settings.SkinVariation.Empty();
}

// Read the settings from the registry
void ReadSettings( StartMenuSettings &settings )
{
	if (g_SettingsDlg)
	{
		ReadSettings(g_SettingsDlg,settings);
		return;
	}
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
	if (regSettings.QueryDWORDValue(L"ScrollMenus",settings.ScrollMenus)!=ERROR_SUCCESS)
		settings.ScrollMenus=DEFAULT_SCROLL_MENUS;
	if (regSettings.QueryDWORDValue(L"Hotkey",settings.Hotkey)!=ERROR_SUCCESS)
		settings.Hotkey=DEFAULT_HOTKEY;
	wchar_t skinName[256];
	ULONG size=_countof(skinName);
	if (regSettings.QueryStringValue(L"SkinName",skinName,&size)==ERROR_SUCCESS)
	{
		settings.SkinName=skinName;
		size=_countof(skinName);
		if (regSettings.QueryStringValue(L"SkinVariation",skinName,&size)==ERROR_SUCCESS)
			settings.SkinVariation=skinName;
	}
	else
	{
		BOOL comp;
		if (!IsAppThemed())
			settings.SkinName=L"Classic Skin";
		else if (LOWORD(GetVersion())==0x0006)
			settings.SkinName=L"Windows Vista Aero";
		else if (SUCCEEDED(DwmIsCompositionEnabled(&comp)) && comp)
			settings.SkinName=L"Windows 7 Aero";
		else
			settings.SkinName=L"Windows 7 Basic";
	}
}

static HRESULT CALLBACK TaskDialogCallbackProc( HWND hwnd, UINT uNotification, WPARAM wParam, LPARAM lParam, LONG_PTR dwRefData )
{
	if (uNotification==TDN_HYPERLINK_CLICKED)
	{
		ShellExecute(hwnd,L"open",(const wchar_t*)lParam,NULL,NULL,SW_SHOWNORMAL);
	}
	return S_OK;
}


static void InitSkinVariations( HWND hwndDlg, const wchar_t *var )
{
	wchar_t skinName[256];
	int idx=(int)SendDlgItemMessage(hwndDlg,IDC_COMBOSKIN,CB_GETCURSEL,0,0);
	SendDlgItemMessage(hwndDlg,IDC_COMBOSKIN,CB_GETLBTEXT,idx,(LPARAM)skinName);
	MenuSkin skin;
	if (!LoadMenuSkin(skinName,skin,NULL,true))
		skin.Variations.clear();

	HWND label=GetDlgItem(hwndDlg,IDC_STATICVAR);
	HWND combo=GetDlgItem(hwndDlg,IDC_COMBOVAR);
	SendMessage(combo,CB_RESETCONTENT,0,0);
	if (skin.Variations.empty())
	{
		ShowWindow(label,SW_HIDE);
		ShowWindow(combo,SW_HIDE);
	}
	else
	{
		ShowWindow(label,SW_SHOW);
		ShowWindow(combo,SW_SHOW);
		idx=0;
		for (int i=0;i<(int)skin.Variations.size();i++)
		{
			SendMessage(combo,CB_ADDSTRING,0,(LPARAM)(const wchar_t*)skin.Variations[i].second);
			if (var && wcscmp(var,skin.Variations[i].second)==0)
				idx=i;
		}
		SendMessage(combo,CB_SETCURSEL,idx,0);
	}
}

// Dialog proc for the settings dialog box. Edits and saves the settings.
static INT_PTR CALLBACK SettingsDlgProc( HWND hwndDlg, UINT uMsg, WPARAM wParam, LPARAM lParam )
{
	static DWORD s_Hotkey;
	if (uMsg==WM_INITDIALOG)
	{
		// get the DLL version. this is a bit hacky. the standard way is to use GetFileVersionInfo and such API.
		// but it takes a file name instead of module handle so it will probably load the DLL a second time.
		// the header of the version resource is a fixed size so we can count on VS_FIXEDFILEINFO to always
		// be at offset 40
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
			swprintf_s(title,L"Settings for Classic Start Menu %d.%d.%d",HIWORD(pVer->dwProductVersionMS),LOWORD(pVer->dwProductVersionMS),HIWORD(pVer->dwProductVersionLS));
		}
		else
			swprintf_s(title,L"Settings for Classic Start Menu");
		SetWindowText(hwndDlg,title);

		HICON icon=(HICON)LoadImage(g_Instance,MAKEINTRESOURCE(IDI_APPICON),IMAGE_ICON,GetSystemMetrics(SM_CXICON),GetSystemMetrics(SM_CYICON),LR_DEFAULTCOLOR);
		SendMessage(hwndDlg,WM_SETICON,ICON_BIG,(LPARAM)icon);
		icon=(HICON)LoadImage(g_Instance,MAKEINTRESOURCE(IDI_APPICON),IMAGE_ICON,GetSystemMetrics(SM_CXSMICON),GetSystemMetrics(SM_CYSMICON),LR_DEFAULTCOLOR);
		SendMessage(hwndDlg,WM_SETICON,ICON_SMALL,(LPARAM)icon);

		StartMenuSettings settings;
		ReadSettings(settings);
		s_Hotkey=settings.Hotkey;
		SetHotkey(s_Hotkey==0?0:1);

		g_SettingsDlg=hwndDlg; // must be after calling ReadSettings

		CheckDlgButton(hwndDlg,IDC_CHECKFAVORITES,settings.ShowFavorites?BST_CHECKED:BST_UNCHECKED);
		CheckDlgButton(hwndDlg,IDC_CHECKDOCUMENTS,settings.ShowDocuments?BST_CHECKED:BST_UNCHECKED);
		CheckDlgButton(hwndDlg,IDC_CHECKLOGOFF,settings.ShowLogOff?BST_CHECKED:BST_UNCHECKED);
		CheckDlgButton(hwndDlg,IDC_CHECKUNDOCK,settings.ShowUndock?BST_CHECKED:BST_UNCHECKED);
		CheckDlgButton(hwndDlg,IDC_CHECKCONTROLPANEL,settings.ExpandControlPanel?BST_CHECKED:BST_UNCHECKED);
		CheckDlgButton(hwndDlg,IDC_CHECKNETWORK,settings.ExpandNetwork?BST_CHECKED:BST_UNCHECKED);
		CheckDlgButton(hwndDlg,IDC_CHECKPRINTERS,settings.ExpandPrinters?BST_CHECKED:BST_UNCHECKED);
		CheckDlgButton(hwndDlg,IDC_CHECKLINKS,settings.ExpandLinks?BST_CHECKED:BST_UNCHECKED);
		CheckDlgButton(hwndDlg,IDC_CHECKSCROLL,settings.ScrollMenus?BST_CHECKED:BST_UNCHECKED);
		CheckDlgButton(hwndDlg,IDC_CHECKHOTKEY,settings.Hotkey!=0?BST_CHECKED:BST_UNCHECKED);
		if (settings.Hotkey!=0 && settings.Hotkey!=1)
			SendDlgItemMessage(hwndDlg,IDC_HOTKEY,HKM_SETHOTKEY,settings.Hotkey,0);
		SendDlgItemMessage(hwndDlg,IDC_HOTKEY,HKM_SETRULES,HKCOMB_NONE|HKCOMB_S|HKCOMB_C,HOTKEYF_CONTROL|HOTKEYF_ALT);
		EnableWindow(GetDlgItem(hwndDlg,IDC_HOTKEY),settings.Hotkey!=0);

		EnableWindow(GetDlgItem(hwndDlg,IDC_CHECKFAVORITES),!SHRestricted(REST_NOFAVORITESMENU));
		EnableWindow(GetDlgItem(hwndDlg,IDC_CHECKUNDOCK),!SHRestricted(REST_NOSMEJECTPC));

		BOOL bDocuments=!SHRestricted(REST_NORECENTDOCSMENU);
		EnableWindow(GetDlgItem(hwndDlg,IDC_CHECKDOCUMENTS),bDocuments);

		DWORD logoff1=SHRestricted(REST_STARTMENULOGOFF);
		DWORD logoff2=SHRestricted(REST_FORCESTARTMENULOGOFF);
		BOOL bLogOff=!SHRestricted(REST_STARTMENULOGOFF) && !SHRestricted(REST_FORCESTARTMENULOGOFF);
		EnableWindow(GetDlgItem(hwndDlg,IDC_CHECKLOGOFF),logoff1==0 && !logoff2);

		bool bNoSetFolders=SHRestricted(REST_NOSETFOLDERS)!=0; // hide control panel, printers, network
		EnableWindow(GetDlgItem(hwndDlg,IDC_CHECKCONTROLPANEL),!bNoSetFolders && !SHRestricted(REST_NOCONTROLPANEL));
		EnableWindow(GetDlgItem(hwndDlg,IDC_CHECKNETWORK),!bNoSetFolders && !SHRestricted(REST_NONETWORKCONNECTIONS));
		EnableWindow(GetDlgItem(hwndDlg,IDC_CHECKPRINTERS),!bNoSetFolders);

		int idx=(int)SendDlgItemMessage(hwndDlg,IDC_COMBOSKIN,CB_ADDSTRING,0,(LPARAM)L"<Default>");
		SendDlgItemMessage(hwndDlg,IDC_COMBOSKIN,CB_SETCURSEL,idx,0);
		wchar_t find[_MAX_PATH];
		GetSkinsPath(find);
		wcscat_s(find,L"1.txt");
		if (GetFileAttributes(find)!=INVALID_FILE_ATTRIBUTES)
		{
			idx=(int)SendDlgItemMessage(hwndDlg,IDC_COMBOSKIN,CB_ADDSTRING,0,(LPARAM)L"Custom");
			if (_wcsicmp(settings.SkinName,L"Custom")==0)
				SendDlgItemMessage(hwndDlg,IDC_COMBOSKIN,CB_SETCURSEL,idx,0);
		}

		*PathFindFileName(find)=0;
		wcscat_s(find,L"*.skin");
		WIN32_FIND_DATA data;
		HANDLE h=FindFirstFile(find,&data);
		while (h!=INVALID_HANDLE_VALUE)
		{
			if (!(data.dwFileAttributes&FILE_ATTRIBUTE_DIRECTORY))
			{
				*PathFindExtension(data.cFileName)=0;
				idx=(int)SendDlgItemMessage(hwndDlg,IDC_COMBOSKIN,CB_ADDSTRING,0,(LPARAM)data.cFileName);
				if (_wcsicmp(settings.SkinName,data.cFileName)==0)
					SendDlgItemMessage(hwndDlg,IDC_COMBOSKIN,CB_SETCURSEL,idx,0);
			}
			if (!FindNextFile(h,&data)) break;
		}
		InitSkinVariations(hwndDlg,settings.SkinVariation);
		return TRUE;
	}

	if (uMsg==WM_COMMAND && wParam==MAKEWPARAM(IDC_COMBOSKIN,CBN_SELENDOK))
	{
		InitSkinVariations(hwndDlg,NULL);
		return TRUE;
	}

	if (uMsg==WM_COMMAND && wParam==IDC_CHECKHOTKEY)
	{
		EnableWindow(GetDlgItem(hwndDlg,IDC_HOTKEY),IsDlgButtonChecked(hwndDlg,IDC_CHECKHOTKEY)==BST_CHECKED);
		return TRUE;
	}

	if (uMsg==WM_COMMAND && wParam==IDC_ABOUT)
	{
		wchar_t name[256];
		int idx=(int)SendDlgItemMessage(hwndDlg,IDC_COMBOSKIN,CB_GETCURSEL,0,0);
		SendDlgItemMessage(hwndDlg,IDC_COMBOSKIN,CB_GETLBTEXT,idx,(LPARAM)name);
		MenuSkin skin;
		wchar_t caption[256];
		swprintf_s(caption,L"About skin %s",name);
		if (!LoadMenuSkin(name,skin,NULL,true))
		{
			MessageBox(hwndDlg,L"Failed to load skin.",caption,MB_OK|MB_ICONERROR);
			return TRUE;
		}
		TASKDIALOGCONFIG task={sizeof(task),hwndDlg,NULL,TDF_ENABLE_HYPERLINKS|TDF_ALLOW_DIALOG_CANCELLATION|TDF_USE_HICON_MAIN,TDCBF_OK_BUTTON};
		task.pszWindowTitle=caption;
		task.pszContent=skin.About;
		task.hMainIcon=skin.AboutIcon?skin.AboutIcon:LoadIcon(NULL,IDI_INFORMATION);
		task.pfCallback=TaskDialogCallbackProc;
		TaskDialogIndirect(&task,NULL,NULL,NULL);
		return TRUE;
	}
	if (uMsg==WM_COMMAND && wParam==IDOK)
	{
		StartMenuSettings settings;
		ReadSettings(hwndDlg,settings);
		CRegKey regSettings;
		if (regSettings.Open(HKEY_CURRENT_USER,L"Software\\IvoSoft\\ClassicStartMenu")!=ERROR_SUCCESS)
			regSettings.Create(HKEY_CURRENT_USER,L"Software\\IvoSoft\\ClassicStartMenu");

		regSettings.SetDWORDValue(L"ShowFavorites",settings.ShowFavorites);
		regSettings.SetDWORDValue(L"ShowDocuments",settings.ShowDocuments);
		regSettings.SetDWORDValue(L"ShowLogOff",settings.ShowLogOff);
		regSettings.SetDWORDValue(L"ShowUndock",settings.ShowUndock);
		regSettings.SetDWORDValue(L"ExpandControlPanel",settings.ExpandControlPanel);
		regSettings.SetDWORDValue(L"ExpandNetwork",settings.ExpandNetwork);
		regSettings.SetDWORDValue(L"ExpandPrinters",settings.ExpandPrinters);
		regSettings.SetDWORDValue(L"ExpandLinks",settings.ExpandLinks);
		regSettings.SetDWORDValue(L"ScrollMenus",settings.ScrollMenus);
		regSettings.SetDWORDValue(L"Hotkey",settings.Hotkey);
		s_Hotkey=settings.Hotkey;
		regSettings.SetStringValue(L"SkinName",settings.SkinName);
		regSettings.SetStringValue(L"SkinVariation",settings.SkinVariation);

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
void EditSettings( bool bModal )
{
	if (g_SettingsDlg)
		SetActiveWindow(g_SettingsDlg);
	else
	{
		HWND dlg=CreateDialog(g_Instance,MAKEINTRESOURCE(IDD_SETTINGS),NULL,SettingsDlgProc);
		ShowWindow(dlg,SW_SHOWNORMAL);
		SetForegroundWindow(dlg);
		if (bModal)
		{
			MSG msg;
			while (g_SettingsDlg && GetMessage(&msg,0,0,0))
			{
				if (IsSettingsMessage(&msg)) continue;
				TranslateMessage(&msg);
				DispatchMessage(&msg);
			}
		}
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
