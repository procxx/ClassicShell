// Classic Shell (c) 2009-2010, Ivo Beltchev
// The sources for Classic Shell are distributed under the MIT open source license

#include "stdafx.h"
#include "resource.h"
#include "Settings.h"
#include "ClassicStartMenuDLL.h"
#include "SkinManager.h"
#include "FNVHash.h"
#include "ParseSettings.h"
#include "dllmain.h"
#include <uxtheme.h>
#include <dwmapi.h>
#include <htmlhelp.h>

// Default settings
const DWORD DEFAULT_SHOW_FAVORITES=0;
const DWORD DEFAULT_SHOW_DOCUMENTS=1;
const DWORD DEFAULT_SHOW_LOGOFF=0;
const DWORD DEFAULT_SHOW_UNDOCK=1;
const DWORD DEFAULT_SHOW_RECENT=0;
const DWORD DEFAULT_EXPAND_CONTROLPANEL=0;
const DWORD DEFAULT_EXPAND_NETWORK=0;
const DWORD DEFAULT_EXPAND_PRINTERS=0;
const DWORD DEFAULT_EXPAND_SHORTCUTS=1;
const DWORD DEFAULT_SMALL_ICONS=0;
const DWORD DEFAULT_THEME=1;
const DWORD DEFAULT_SCROLL_MENUS=0;
const DWORD DEFAULT_REPORT_ERRORS=0;
const DWORD DEFAULT_CONTROLS=(StartMenuSettings::OPEN_CLASSIC*0x00010001)|(StartMenuSettings::OPEN_WINDOWS*0x01000100); // Click and Win open CSM, Shift+Click and Shift+Win open NSM

static HWND g_SettingsDlg;
static int g_VariationY, g_NoVariationY;
static std::vector<unsigned int> g_ExtraOptions; // options that were set but are not used by the current skin

enum
{
	CHECK_FALSE,
	CHECK_TRUE,
	CHECK_FALSE_DISABLED,
	CHECK_TRUE_DISABLED,
};

struct ListOption
{
	CString name;
	CString condition;
	unsigned int hash;
	bool bValue; // value set by the user
	bool bValue2; // default value when the condition is false
	bool bFinalValue; // final value used by the skin
	bool bDisabled; // when the condition is false
};

static std::vector<ListOption> g_ListOptions;

static void ReadSettings( HWND hwndDlg, StartMenuSettings &settings )
{
	settings.ShowFavorites=(IsDlgButtonChecked(hwndDlg,IDC_CHECKFAVORITES)==BST_CHECKED)?1:0;
	settings.ShowDocuments=(IsDlgButtonChecked(hwndDlg,IDC_CHECKDOCUMENTS)==BST_CHECKED)?1:0;
	settings.ShowLogOff=(IsDlgButtonChecked(hwndDlg,IDC_CHECKLOGOFF)==BST_CHECKED)?1:0;
	settings.ShowUndock=(IsDlgButtonChecked(hwndDlg,IDC_CHECKUNDOCK)==BST_CHECKED)?1:0;
	settings.ShowRecent=(IsDlgButtonChecked(hwndDlg,IDC_CHECKRECENT)==BST_CHECKED)?1:0;
	settings.ExpandControlPanel=(IsDlgButtonChecked(hwndDlg,IDC_CHECKCONTROLPANEL)==BST_CHECKED)?1:0;
	settings.ExpandNetwork=(IsDlgButtonChecked(hwndDlg,IDC_CHECKNETWORK)==BST_CHECKED)?1:0;
	settings.ExpandPrinters=(IsDlgButtonChecked(hwndDlg,IDC_CHECKPRINTERS)==BST_CHECKED)?1:0;
	settings.ExpandLinks=(IsDlgButtonChecked(hwndDlg,IDC_CHECKLINKS)==BST_CHECKED)?1:0;
	settings.ScrollMenus=(DWORD)SendDlgItemMessage(hwndDlg,IDC_COMBOSCROLL,CB_GETCURSEL,0,0);
	settings.ReportErrors=(IsDlgButtonChecked(hwndDlg,IDC_CHECKERRORS)==BST_CHECKED)?1:0;
	settings.HotkeyCSM=(DWORD)SendDlgItemMessage(hwndDlg,IDC_HOTKEY,HKM_GETHOTKEY,0,0);
	settings.HotkeyNSM=(DWORD)SendDlgItemMessage(hwndDlg,IDC_HOTKEYW,HKM_GETHOTKEY,0,0);
	settings.Controls=(DWORD)SendDlgItemMessage(hwndDlg,IDC_COMBOCLICK,CB_GETCURSEL,0,0);
	settings.Controls|=(DWORD)SendDlgItemMessage(hwndDlg,IDC_COMBOSCLICK,CB_GETCURSEL,0,0)<<8;
	settings.Controls|=(DWORD)SendDlgItemMessage(hwndDlg,IDC_COMBOWIN,CB_GETCURSEL,0,0)<<16;
	settings.Controls|=(DWORD)SendDlgItemMessage(hwndDlg,IDC_COMBOSWIN,CB_GETCURSEL,0,0)<<24;
	settings.Controls|=(DWORD)SendDlgItemMessage(hwndDlg,IDC_COMBOMCLICK,CB_GETCURSEL,0,0)<<4;
	settings.Controls|=(DWORD)SendDlgItemMessage(hwndDlg,IDC_COMBOHOVER,CB_GETCURSEL,0,0)<<12;
	if (IsDlgButtonChecked(hwndDlg,IDC_CHECKEXPANDPROGRAMS)==BST_CHECKED)
	{
		settings.Controls|=0x10000000;
		if (IsDlgButtonChecked(hwndDlg,IDC_RADIOALL)==BST_CHECKED)
			settings.Controls|=0x20000000;
	}

	wchar_t skinName[256];
	int idx=(int)SendDlgItemMessage(hwndDlg,IDC_COMBOSKIN,CB_GETCURSEL,0,0);
	SendDlgItemMessage(hwndDlg,IDC_COMBOSKIN,CB_GETLBTEXT,idx,(LPARAM)skinName);
	if (wcscmp(skinName,LoadStringEx(IDS_DEFAULT_SKIN))==0)
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

	HWND list=GetDlgItem(hwndDlg,IDC_LISTOPTIONS);
	int n=(int)g_ListOptions.size();
	settings.SkinOptions.resize(n);
	for (int i=0;i<n;i++)
		settings.SkinOptions[i]=(g_ListOptions[i].hash|(g_ListOptions[i].bValue?1:0));

	idx=(int)SendDlgItemMessage(hwndDlg,IDC_COMBOSKIN2,CB_GETCURSEL,0,0);
	SendDlgItemMessage(hwndDlg,IDC_COMBOSKIN2,CB_GETLBTEXT,idx,(LPARAM)skinName);
	if (wcscmp(skinName,LoadStringEx(IDS_DEFAULT_SKIN))==0)
		settings.SkinName2.Empty();
	else if (wcscmp(skinName,LoadStringEx(IDS_MAIN_SKIN))==0)
		settings.SkinName2=L"<Main Skin>";
	else
		settings.SkinName2=skinName;
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
	if (regSettings.QueryDWORDValue(L"ShowRecent",settings.ShowRecent)!=ERROR_SUCCESS)
		settings.ShowRecent=DEFAULT_SHOW_RECENT;
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
	if (settings.ScrollMenus<0) settings.ScrollMenus=0;
	if (settings.ScrollMenus>2) settings.ScrollMenus=2;
	if (regSettings.QueryDWORDValue(L"ReportErrors",settings.ReportErrors)!=ERROR_SUCCESS)
		settings.ReportErrors=DEFAULT_REPORT_ERRORS;
	if (regSettings.QueryDWORDValue(L"HotkeyCSM",settings.HotkeyCSM)!=ERROR_SUCCESS)
		settings.HotkeyCSM=0;
	if (regSettings.QueryDWORDValue(L"HotkeyNSM",settings.HotkeyNSM)!=ERROR_SUCCESS)
		settings.HotkeyNSM=0;
	if (regSettings.QueryDWORDValue(L"Controls",settings.Controls)!=ERROR_SUCCESS)
		settings.Controls=DEFAULT_CONTROLS;
	wchar_t skinName[256];
	ULONG size=_countof(skinName);
	settings.SkinOptions.clear();
	if (regSettings.QueryStringValue(L"SkinName",skinName,&size)==ERROR_SUCCESS)
	{
		settings.SkinName=skinName;
		size=_countof(skinName);
		if (regSettings.QueryStringValue(L"SkinVariation",skinName,&size)==ERROR_SUCCESS)
			settings.SkinVariation=skinName;

		size=0;
		regSettings.QueryBinaryValue(L"SkinOptions",NULL,&size);
		if (size>0)
		{
			settings.SkinOptions.resize(size/4);
			regSettings.QueryBinaryValue(L"SkinOptions",&settings.SkinOptions[0],&size);
		}
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

	size=_countof(skinName);
	if (regSettings.QueryStringValue(L"SkinName2",skinName,&size)==ERROR_SUCCESS)
		settings.SkinName2=skinName;
	else
		settings.SkinName2.Empty();
}

static HRESULT CALLBACK TaskDialogCallbackProc( HWND hwnd, UINT uNotification, WPARAM wParam, LPARAM lParam, LONG_PTR dwRefData )
{
	if (uNotification==TDN_HYPERLINK_CLICKED)
	{
		ShellExecute(hwnd,L"open",(const wchar_t*)lParam,NULL,NULL,SW_SHOWNORMAL);
	}
	return S_OK;
}


static void InitSkinVariations( HWND hwndDlg, const wchar_t *var, const std::vector<unsigned int> &options )
{
	wchar_t skinName[256];
	int idx=(int)SendDlgItemMessage(hwndDlg,IDC_COMBOSKIN,CB_GETCURSEL,0,0);
	SendDlgItemMessage(hwndDlg,IDC_COMBOSKIN,CB_GETLBTEXT,idx,(LPARAM)skinName);
	MenuSkin skin;
	if (!LoadMenuSkin(skinName,skin,NULL,options,RES_LEVEL_NONE))
	{
		skin.Variations.clear();
		skin.Options.clear();
	}

	ShowWindow(GetDlgItem(hwndDlg,IDC_STATICVER),skin.version>MAX_SKIN_VERSION);

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

	label=GetDlgItem(hwndDlg,IDC_STATICOPT);
	HWND list=GetDlgItem(hwndDlg,IDC_LISTOPTIONS);

	// store current options in g_ExtraOptions and clear the list
	int n=(int)g_ListOptions.size();
	for (int i=0;i<n;i++)
	{
		unsigned int hash=g_ListOptions[i].hash;
		bool bValue=g_ListOptions[i].bValue;
		bool bFound=false;
		for (std::vector<unsigned int>::iterator it=g_ExtraOptions.begin();it!=g_ExtraOptions.end();++it)
			if ((*it&0xFFFFFFFE)==hash)
			{
				*it=hash|(bValue?1:0);
				bFound=true;
				break;
			}
		if (!bFound)
			g_ExtraOptions.push_back(hash|(bValue?1:0));
	}
	SendMessage(list,LVM_DELETEALLITEMS,0,0);

	if (skin.Options.empty())
	{
		ShowWindow(label,SW_HIDE);
		ShowWindow(list,SW_HIDE);
	}
	else
	{
		RECT rc;
		GetWindowRect(label,&rc);
		MapWindowPoints(NULL,hwndDlg,(POINT*)&rc,2);
		int dy=(skin.Variations.empty()?g_NoVariationY:g_VariationY)-rc.top;
		SetWindowPos(label,NULL,rc.left,rc.top+dy,0,0,SWP_NOZORDER|SWP_NOSIZE|SWP_SHOWWINDOW);

		GetWindowRect(list,&rc);
		MapWindowPoints(NULL,hwndDlg,(POINT*)&rc,2);
		SetWindowPos(list,NULL,rc.left,rc.top+dy,rc.right-rc.left,rc.bottom-dy-rc.top,SWP_NOZORDER|SWP_SHOWWINDOW);

		int n=(int)skin.Options.size();
		g_ListOptions.clear();
		g_ListOptions.resize(n);
		std::vector<const wchar_t*> values; // list of true values
		for (int i=0;i<n;i++)
		{
			LVITEM item={LVIF_TEXT|LVIF_STATE,i};
			item.pszText=(LPWSTR)(LPCWSTR)skin.Options[i].label;
			unsigned int hash=CalcFNVHash(skin.Options[i].name)&0xFFFFFFFE;
			g_ListOptions[i].name=skin.Options[i].name;
			g_ListOptions[i].hash=hash;
			g_ListOptions[i].condition=skin.Options[i].condition;

			// get the default value
			bool bValue=skin.Options[i].value;

			// override from g_ExtraOptions
			for (std::vector<unsigned int>::const_iterator it=g_ExtraOptions.begin();it!=g_ExtraOptions.end();++it)
				if ((*it&0xFFFFFFFE)==hash)
				{
					bValue=(*it&1)==1;
					break;
				}

			// override from options
			for (std::vector<unsigned int>::const_iterator it=options.begin();it!=options.end();++it)
				if ((*it&0xFFFFFFFE)==hash)
				{
					bValue=(*it&1)==1;
					break;
				}

			g_ListOptions[i].bValue=bValue;
			g_ListOptions[i].bValue2=skin.Options[i].value2;

			bool bDisabled=false;
			if (!skin.Options[i].condition.IsEmpty() && !EvalCondition(skin.Options[i].condition,values.empty()?NULL:&values[0],(int)values.size()))
			{
				bDisabled=true;
				bValue=skin.Options[i].value2;
			}
			g_ListOptions[i].bDisabled=bDisabled;

			if (bValue)
				values.push_back(g_ListOptions[i].name);
			g_ListOptions[i].bFinalValue=bValue;

			int state=bDisabled?(bValue?CHECK_TRUE_DISABLED:CHECK_FALSE_DISABLED):(bValue?CHECK_TRUE:CHECK_FALSE);
			item.state=INDEXTOSTATEIMAGEMASK(state+1);
			item.stateMask=LVIS_STATEIMAGEMASK;

			SendMessage(list,LVM_INSERTITEM,0,(LPARAM)&item);
		}
	}
}

static void ToggleListOption( HWND list, int idx )
{
	if (g_ListOptions[idx].bDisabled) return;
	g_ListOptions[idx].bValue=!g_ListOptions[idx].bValue;
	ListView_SetItemState(list,idx,INDEXTOSTATEIMAGEMASK((g_ListOptions[idx].bValue?CHECK_TRUE:CHECK_FALSE)+1),LVIS_STATEIMAGEMASK);

	std::vector<const wchar_t*> values; // list of true values
	int n=(int)g_ListOptions.size();
	for (int i=0;i<n;i++)
	{
		bool bValue=g_ListOptions[i].bValue;
		bool bDisabled=false;
		if (!g_ListOptions[i].condition.IsEmpty() && !EvalCondition(g_ListOptions[i].condition,values.empty()?NULL:&values[0],(int)values.size()))
		{
			bDisabled=true;
			bValue=g_ListOptions[i].bValue2;
		}
		if (bValue)
			values.push_back(g_ListOptions[i].name);

		if (bValue!=g_ListOptions[i].bFinalValue || bDisabled!=g_ListOptions[i].bDisabled)
		{
			g_ListOptions[i].bFinalValue=bValue;
			g_ListOptions[i].bDisabled=bDisabled;
			int state=bDisabled?(bValue?CHECK_TRUE_DISABLED:CHECK_FALSE_DISABLED):(bValue?CHECK_TRUE:CHECK_FALSE);
			ListView_SetItemState(list,i,INDEXTOSTATEIMAGEMASK(state+1),LVIS_STATEIMAGEMASK);
			RECT rc;
			ListView_GetItemRect(list,i,&rc,LVIR_BOUNDS);
			InvalidateRect(list,&rc,TRUE);
		}
	}
}

// Dialog proc for the settings dialog box. Edits and saves the settings.
static INT_PTR CALLBACK SettingsDlgProc( HWND hwndDlg, UINT uMsg, WPARAM wParam, LPARAM lParam )
{
	static DWORD s_HotkeyCSM, s_HotkeyNSM, s_Controls;
	static HICON s_ShieldIcon;
	if (uMsg==WM_INITDIALOG)
	{
		DWORD ver=GetVersionEx(g_Instance);
		wchar_t title[100];
		if (ver)
			Sprintf(title,_countof(title),LoadStringEx(IDS_SETTINGS_TITLE_VER),ver>>24,(ver>>16)&0xFF,ver&0xFFFF);
		else
			Sprintf(title,_countof(title),LoadStringEx(IDS_SETTINGS_TITLE));
		SetWindowText(hwndDlg,title);

		HICON icon=(HICON)LoadImage(g_Instance,MAKEINTRESOURCE(IDI_APPICON),IMAGE_ICON,GetSystemMetrics(SM_CXICON),GetSystemMetrics(SM_CYICON),LR_DEFAULTCOLOR);
		SendMessage(hwndDlg,WM_SETICON,ICON_BIG,(LPARAM)icon);
		icon=(HICON)LoadImage(g_Instance,MAKEINTRESOURCE(IDI_APPICON),IMAGE_ICON,GetSystemMetrics(SM_CXSMICON),GetSystemMetrics(SM_CYSMICON),LR_DEFAULTCOLOR);
		SendMessage(hwndDlg,WM_SETICON,ICON_SMALL,(LPARAM)icon);

		StartMenuSettings settings;
		ReadSettings(settings);
		s_HotkeyCSM=settings.HotkeyCSM;
		s_HotkeyNSM=settings.HotkeyNSM;
		s_Controls=settings.Controls;
		SetControls(0,0,settings.Controls);
		g_ExtraOptions.clear();

		g_SettingsDlg=hwndDlg; // must be after calling ReadSettings

		SendDlgItemMessage(hwndDlg,IDC_COMBOSCROLL,CB_ADDSTRING,0,(LPARAM)(const wchar_t*)LoadStringEx(IDS_SCROLL_NO));
		SendDlgItemMessage(hwndDlg,IDC_COMBOSCROLL,CB_ADDSTRING,0,(LPARAM)(const wchar_t*)LoadStringEx(IDS_SCROLL_YES));
		SendDlgItemMessage(hwndDlg,IDC_COMBOSCROLL,CB_ADDSTRING,0,(LPARAM)(const wchar_t*)LoadStringEx(IDS_SCROLL_AUTO));

		CheckDlgButton(hwndDlg,IDC_CHECKFAVORITES,settings.ShowFavorites?BST_CHECKED:BST_UNCHECKED);
		CheckDlgButton(hwndDlg,IDC_CHECKDOCUMENTS,settings.ShowDocuments?BST_CHECKED:BST_UNCHECKED);
		CheckDlgButton(hwndDlg,IDC_CHECKLOGOFF,settings.ShowLogOff?BST_CHECKED:BST_UNCHECKED);
		CheckDlgButton(hwndDlg,IDC_CHECKUNDOCK,settings.ShowUndock?BST_CHECKED:BST_UNCHECKED);
		CheckDlgButton(hwndDlg,IDC_CHECKRECENT,settings.ShowRecent?BST_CHECKED:BST_UNCHECKED);
		CheckDlgButton(hwndDlg,IDC_CHECKCONTROLPANEL,settings.ExpandControlPanel?BST_CHECKED:BST_UNCHECKED);
		CheckDlgButton(hwndDlg,IDC_CHECKNETWORK,settings.ExpandNetwork?BST_CHECKED:BST_UNCHECKED);
		CheckDlgButton(hwndDlg,IDC_CHECKPRINTERS,settings.ExpandPrinters?BST_CHECKED:BST_UNCHECKED);
		CheckDlgButton(hwndDlg,IDC_CHECKLINKS,settings.ExpandLinks?BST_CHECKED:BST_UNCHECKED);
		SendDlgItemMessage(hwndDlg,IDC_COMBOSCROLL,CB_SETCURSEL,settings.ScrollMenus,0);
		CheckDlgButton(hwndDlg,IDC_CHECKERRORS,settings.ReportErrors?BST_CHECKED:BST_UNCHECKED);
		SendDlgItemMessage(hwndDlg,IDC_HOTKEY,HKM_SETHOTKEY,settings.HotkeyCSM,0);
		SendDlgItemMessage(hwndDlg,IDC_HOTKEYW,HKM_SETHOTKEY,settings.HotkeyNSM,0);
		SendDlgItemMessage(hwndDlg,IDC_HOTKEY,HKM_SETRULES,HKCOMB_NONE|HKCOMB_S|HKCOMB_C,HOTKEYF_CONTROL|HOTKEYF_ALT);
		SendDlgItemMessage(hwndDlg,IDC_HOTKEYW,HKM_SETRULES,HKCOMB_NONE|HKCOMB_S|HKCOMB_C,HOTKEYF_CONTROL|HOTKEYF_ALT);

		CString strCtrl0=LoadStringEx(IDS_CTRL_NOTHING);
		CString strCtrl1=LoadStringEx(IDS_CTRL_CLASSIC);
		CString strCtrl2=LoadStringEx(IDS_CTRL_WINDOWS);

		const wchar_t *text[]={strCtrl0,strCtrl1,strCtrl2};
		for (int i=0;i<_countof(text);i++)
		{
			SendDlgItemMessage(hwndDlg,IDC_COMBOCLICK,CB_ADDSTRING,0,(LPARAM)text[i]);
			SendDlgItemMessage(hwndDlg,IDC_COMBOSCLICK,CB_ADDSTRING,0,(LPARAM)text[i]);
			SendDlgItemMessage(hwndDlg,IDC_COMBOWIN,CB_ADDSTRING,0,(LPARAM)text[i]);
			SendDlgItemMessage(hwndDlg,IDC_COMBOSWIN,CB_ADDSTRING,0,(LPARAM)text[i]);
			SendDlgItemMessage(hwndDlg,IDC_COMBOMCLICK,CB_ADDSTRING,0,(LPARAM)text[i]);
			SendDlgItemMessage(hwndDlg,IDC_COMBOHOVER,CB_ADDSTRING,0,(LPARAM)text[i]);
		}
		SendDlgItemMessage(hwndDlg,IDC_COMBOCLICK,CB_SETCURSEL,settings.Controls&15,0);
		SendDlgItemMessage(hwndDlg,IDC_COMBOSCLICK,CB_SETCURSEL,(settings.Controls>>8)&15,0);
		SendDlgItemMessage(hwndDlg,IDC_COMBOWIN,CB_SETCURSEL,(settings.Controls>>16)&15,0);
		SendDlgItemMessage(hwndDlg,IDC_COMBOSWIN,CB_SETCURSEL,(settings.Controls>>24)&15,0);
		SendDlgItemMessage(hwndDlg,IDC_COMBOMCLICK,CB_SETCURSEL,(settings.Controls>>4)&15,0);
		SendDlgItemMessage(hwndDlg,IDC_COMBOHOVER,CB_SETCURSEL,(settings.Controls>>12)&15,0);
		CheckDlgButton(hwndDlg,IDC_CHECKEXPANDPROGRAMS,(settings.Controls&0x10000000)?BST_CHECKED:BST_UNCHECKED);
		CheckDlgButton(hwndDlg,IDC_RADIOSEARCH,(settings.Controls&0x20000000)?BST_UNCHECKED:BST_CHECKED);
		CheckDlgButton(hwndDlg,IDC_RADIOALL,(settings.Controls&0x20000000)?BST_CHECKED:BST_UNCHECKED);

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

		// create state images for the options list
		int size=14;
		RECT rc={0,0,size,size};
		HIMAGELIST images=ImageList_Create(size,size,ILC_COLOR32,0,4);
		BITMAPINFO dib={sizeof(dib)};
		dib.bmiHeader.biWidth=size;
		dib.bmiHeader.biHeight=-size;
		dib.bmiHeader.biPlanes=1;
		dib.bmiHeader.biBitCount=32;
		dib.bmiHeader.biCompression=BI_RGB;
		HDC hdc=CreateCompatibleDC(NULL);
		HBITMAP bmp=CreateDIBSection(hdc,&dib,DIB_RGB_COLORS,NULL,NULL,0);
		HBITMAP bmp0;

		for (int i=0;i<4;i++)
		{
			bmp0=(HBITMAP)SelectObject(hdc,bmp);
			FillRect(hdc,&rc,(HBRUSH)(COLOR_WINDOW+1));
			UINT state=DFCS_BUTTONCHECK|DFCS_FLAT;
			if (i&1) state|=DFCS_CHECKED;
			if (i&2) state|=DFCS_INACTIVE;
			DrawFrameControl(hdc,&rc,DFC_BUTTON,state);
			SelectObject(hdc,bmp0);
			ImageList_Add(images,bmp,NULL);
		}

		DeleteObject(bmp);
		DeleteDC(hdc);

		GetClientRect(GetDlgItem(hwndDlg,IDC_LISTOPTIONS),&rc);
		LVCOLUMN column={LVCF_FMT|LVCF_WIDTH,LVCFMT_LEFT,rc.right-GetSystemMetrics(SM_CXVSCROLL)};
		SendDlgItemMessage(hwndDlg,IDC_LISTOPTIONS,LVM_INSERTCOLUMN,0,(LPARAM)&column);
	
		SendDlgItemMessage(hwndDlg,IDC_LISTOPTIONS,LVM_SETIMAGELIST,LVSIL_STATE,(LPARAM)images);

		POINT pt={0,0};
		GetWindowRect(GetDlgItem(hwndDlg,IDC_STATICOPT),&rc);
		pt.y=rc.top;
		ScreenToClient(hwndDlg,&pt);
		g_VariationY=pt.y;
		GetWindowRect(GetDlgItem(hwndDlg,IDC_STATICVAR),&rc);
		pt.y=rc.top;
		ScreenToClient(hwndDlg,&pt);
		g_NoVariationY=pt.y;

		int idx=(int)SendDlgItemMessage(hwndDlg,IDC_COMBOSKIN,CB_ADDSTRING,0,(LPARAM)(const wchar_t*)LoadStringEx(IDS_DEFAULT_SKIN));
		SendDlgItemMessage(hwndDlg,IDC_COMBOSKIN,CB_SETCURSEL,idx,0);
		int idx2=(int)SendDlgItemMessage(hwndDlg,IDC_COMBOSKIN2,CB_ADDSTRING,0,(LPARAM)(const wchar_t*)LoadStringEx(IDS_DEFAULT_SKIN));
		SendDlgItemMessage(hwndDlg,IDC_COMBOSKIN2,CB_SETCURSEL,idx2,0);
		idx2=(int)SendDlgItemMessage(hwndDlg,IDC_COMBOSKIN2,CB_ADDSTRING,0,(LPARAM)(const wchar_t*)LoadStringEx(IDS_MAIN_SKIN));
		if (_wcsicmp(settings.SkinName2,L"<Main Skin>")==0)
			SendDlgItemMessage(hwndDlg,IDC_COMBOSKIN2,CB_SETCURSEL,idx2,0);
		wchar_t find[_MAX_PATH];
		GetSkinsPath(find);
		Strcat(find,_countof(find),L"1.txt");
		if (GetFileAttributes(find)!=INVALID_FILE_ATTRIBUTES)
		{
			idx=(int)SendDlgItemMessage(hwndDlg,IDC_COMBOSKIN,CB_ADDSTRING,0,(LPARAM)L"Custom");
			if (_wcsicmp(settings.SkinName,L"Custom")==0)
				SendDlgItemMessage(hwndDlg,IDC_COMBOSKIN,CB_SETCURSEL,idx,0);
		}

		*PathFindFileName(find)=0;
		Strcat(find,_countof(find),L"*.skin");
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

				idx2=(int)SendDlgItemMessage(hwndDlg,IDC_COMBOSKIN2,CB_ADDSTRING,0,(LPARAM)data.cFileName);
				if (_wcsicmp(settings.SkinName2,data.cFileName)==0)
					SendDlgItemMessage(hwndDlg,IDC_COMBOSKIN2,CB_SETCURSEL,idx2,0);
			}
			if (!FindNextFile(h,&data))
			{
				FindClose(h);
				break;
			}
		}
		InitSkinVariations(hwndDlg,settings.SkinVariation,settings.SkinOptions);

		SHSTOCKICONINFO sii={sizeof(sii)};
		SHGetStockIconInfo(SIID_SHIELD,SHGSI_ICON|SHGSI_SMALLICON,&sii);
		s_ShieldIcon=sii.hIcon;
		HWND shield=GetDlgItem(hwndDlg,IDC_SHIELD);
		GetWindowRect(shield,&rc);
		pt.x=rc.right;
		pt.y=rc.bottom;
		ScreenToClient(hwndDlg,&pt);
		int iconSize=GetSystemMetrics(SM_CXSMICON);
		SetWindowPos(shield,NULL,pt.x-iconSize,pt.y-iconSize,iconSize,iconSize,SWP_NOZORDER);
		SendMessage(shield,STM_SETICON,(WPARAM)sii.hIcon,0);
	}

	if (uMsg==WM_INITDIALOG || (uMsg==WM_COMMAND && wParam==IDC_CHECKEXPANDPROGRAMS))
	{
		BOOL enabled=(IsDlgButtonChecked(hwndDlg,IDC_CHECKEXPANDPROGRAMS)==BST_CHECKED);
		EnableWindow(GetDlgItem(hwndDlg,IDC_STATICSKIN2),enabled);
		EnableWindow(GetDlgItem(hwndDlg,IDC_COMBOSKIN2),enabled);
		EnableWindow(GetDlgItem(hwndDlg,IDC_STATICFOCUS),enabled);
		EnableWindow(GetDlgItem(hwndDlg,IDC_RADIOSEARCH),enabled);
		EnableWindow(GetDlgItem(hwndDlg,IDC_RADIOALL),enabled);
		return TRUE;
	}

	if (uMsg==WM_COMMAND && wParam==MAKEWPARAM(IDC_COMBOSKIN,CBN_SELENDOK))
	{
		InitSkinVariations(hwndDlg,NULL,std::vector<unsigned int>());
		return TRUE;
	}

	if (uMsg==WM_COMMAND && wParam==IDC_ABOUT)
	{
		wchar_t name[256];
		int idx=(int)SendDlgItemMessage(hwndDlg,IDC_COMBOSKIN,CB_GETCURSEL,0,0);
		SendDlgItemMessage(hwndDlg,IDC_COMBOSKIN,CB_GETLBTEXT,idx,(LPARAM)name);
		MenuSkin skin;
		wchar_t caption[256];
		Sprintf(caption,_countof(caption),LoadStringEx(IDS_SKIN_ABOUT),name);
		if (!LoadMenuSkin(name,skin,NULL,std::vector<unsigned int>(),RES_LEVEL_NONE))
		{
			MessageBox(hwndDlg,LoadStringEx(IDS_SKIN_FAIL),caption,MB_OK|MB_ICONERROR);
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
		regSettings.SetDWORDValue(L"ShowRecent",settings.ShowRecent);
		regSettings.SetDWORDValue(L"ExpandControlPanel",settings.ExpandControlPanel);
		regSettings.SetDWORDValue(L"ExpandNetwork",settings.ExpandNetwork);
		regSettings.SetDWORDValue(L"ExpandPrinters",settings.ExpandPrinters);
		regSettings.SetDWORDValue(L"ExpandLinks",settings.ExpandLinks);
		regSettings.SetDWORDValue(L"ScrollMenus",settings.ScrollMenus);
		regSettings.SetDWORDValue(L"ReportErrors",settings.ReportErrors);
		regSettings.SetDWORDValue(L"HotkeyCSM",settings.HotkeyCSM);
		regSettings.SetDWORDValue(L"HotkeyNSM",settings.HotkeyNSM);
		regSettings.SetDWORDValue(L"Controls",settings.Controls);
		s_HotkeyCSM=settings.HotkeyCSM;
		s_HotkeyNSM=settings.HotkeyNSM;
		s_Controls=settings.Controls;
		regSettings.SetStringValue(L"SkinName",settings.SkinName);
		regSettings.SetStringValue(L"SkinVariation",settings.SkinVariation);
		if (settings.SkinOptions.empty())
			regSettings.SetBinaryValue(L"SkinOptions",NULL,0);
		else
			regSettings.SetBinaryValue(L"SkinOptions",&settings.SkinOptions[0],(int)settings.SkinOptions.size()*4);
		regSettings.SetStringValue(L"SkinName2",settings.SkinName2);

		if (!settings.ShowRecent)
			regSettings.DeleteSubKey(L"MRU");

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
			// click on Help
			wchar_t path[_MAX_PATH];
			GetModuleFileName(g_Instance,path,_countof(path));
			*PathFindFileName(path)=0;
			Strcat(path,_countof(path),DOC_PATH L"ClassicShell.chm::/ClassicStartMenu.html");
			HtmlHelp(GetDesktopWindow(),path,HH_DISPLAY_TOPIC,NULL);
			return TRUE;
		}
		if (pHdr->idFrom==IDC_LINKINI && (pHdr->code==NM_CLICK || pHdr->code==NM_RETURN))
		{
			// click on More Settings
			CRegKey regSettings;
			if (regSettings.Open(HKEY_CURRENT_USER,L"Software\\IvoSoft\\ClassicStartMenu")!=ERROR_SUCCESS)
				regSettings.Create(HKEY_CURRENT_USER,L"Software\\IvoSoft\\ClassicStartMenu");

			DWORD show;
			if (regSettings.QueryDWORDValue(L"IgnoreIniWarning",show)!=ERROR_SUCCESS)
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

				TaskDialogIndirect(&task,NULL,NULL,&bIgnore);
				if (bIgnore)
					regSettings.SetDWORDValue(L"IgnoreIniWarning",1);
			}

			// run Notepad as administrator (the ini file is most likely in a protected folder)
			wchar_t path[_MAX_PATH];
			GetModuleFileName(g_Instance,path,_countof(path));
			*PathFindFileName(path)=0;
			Strcat(path,_countof(path),INI_PATH L"StartMenu.ini");
			ShellExecute(hwndDlg,L"runas",L"notepad",path,NULL,SW_SHOWNORMAL);
			return TRUE;
		}
		if (pHdr->idFrom==IDC_LISTOPTIONS && (pHdr->code==NM_CLICK || pHdr->code==NM_DBLCLK))
		{
			// click on the state image
			NMITEMACTIVATE *pItem=(NMITEMACTIVATE*)pHdr;
			LVHITTESTINFO hit={pItem->ptAction};
			int idx=ListView_HitTest(pHdr->hwndFrom,&hit);
			if (idx>=0 && (hit.flags&LVHT_ONITEMSTATEICON))
				ToggleListOption(pHdr->hwndFrom,idx);
		}
		if (pHdr->idFrom==IDC_LISTOPTIONS && pHdr->code==LVN_KEYDOWN && ((NMLVKEYDOWN*)pHdr)->wVKey==VK_SPACE)
		{
			// spacebar was pressed
			int idx=ListView_GetNextItem(pHdr->hwndFrom,-1,LVNI_SELECTED);
			if (idx>=0)
				ToggleListOption(pHdr->hwndFrom,idx);
		}
		if (pHdr->idFrom==IDC_LISTOPTIONS && pHdr->code==NM_CUSTOMDRAW)
		{
			// gray out the disabled options
			NMLVCUSTOMDRAW *pDraw=(NMLVCUSTOMDRAW*)pHdr;
			if (pDraw->nmcd.dwDrawStage==CDDS_PREPAINT)
			{
				SetWindowLongPtr(hwndDlg,DWLP_MSGRESULT,CDRF_NOTIFYITEMDRAW);
				return TRUE;
			}
			if (pDraw->nmcd.dwDrawStage==CDDS_ITEMPREPAINT && g_ListOptions[pDraw->nmcd.dwItemSpec].bDisabled)
			{
				pDraw->clrText=GetSysColor(COLOR_GRAYTEXT);
				SetWindowLongPtr(hwndDlg,DWLP_MSGRESULT,CDRF_DODEFAULT);
				return TRUE;
			}
		}
	}
	if (uMsg==WM_DESTROY)
	{
		SetControls(s_HotkeyCSM,s_HotkeyNSM,s_Controls);
		g_SettingsDlg=NULL;
		if (s_ShieldIcon) DestroyIcon(s_ShieldIcon);
		s_ShieldIcon=NULL;
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
		HWND dlg=CreateSettingsDialog(NULL,SettingsDlgProc);
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
