// Classic Shell (c) 2009-2010, Ivo Beltchev
// The sources for Classic Shell are distributed under the MIT open source license

#include "stdafx.h"
#include "resource.h"
#include "Settings.h"
#include "ClassicStartMenuDLL.h"
#include "SkinManager.h"
#include "FNVHash.h"
#include <uxtheme.h>
#include <dwmapi.h>

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

	HWND list=GetDlgItem(hwndDlg,IDC_LISTOPTIONS);
	int n=(int)SendMessage(list,LVM_GETITEMCOUNT,0,0);
	settings.SkinOptions.resize(n);
	for (int i=0;i<n;i++)
	{
		LVITEM item={LVIF_PARAM,i};
		SendMessage(list,LVM_GETITEM,0,(LPARAM)&item);
		settings.SkinOptions[i]=(unsigned int)(item.lParam|(ListView_GetCheckState(list,i)?1:0));
	}
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
	if (!LoadMenuSkin(skinName,skin,NULL,options,true))
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

	int n=(int)SendMessage(list,LVM_GETITEMCOUNT,0,0);
	for (int i=0;i<n;i++)
	{
		LVITEM item={LVIF_PARAM,i};
		SendMessage(list,LVM_GETITEM,0,(LPARAM)&item);
		unsigned int hash=(unsigned int)item.lParam&0xFFFFFFFE;
		bool bFound=false;
		bool bValue=ListView_GetCheckState(list,i)!=0;
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

		idx=0;
		for (int i=0;i<(int)skin.Options.size();i++)
		{
			LVITEM item={LVIF_PARAM|LVIF_TEXT,i};
			item.pszText=(LPWSTR)(LPCWSTR)skin.Options[i].label;
			unsigned int hash=CalcFNVHash(skin.Options[i].name)&0xFFFFFFFE;

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

			item.lParam=hash;
			SendMessage(list,LVM_INSERTITEM,0,(LPARAM)&item);
			if (bValue)
				ListView_SetCheckState(list,i,TRUE);
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
			Sprintf(title,_countof(title),L"Settings for Classic Start Menu %d.%d.%d",HIWORD(pVer->dwProductVersionMS),LOWORD(pVer->dwProductVersionMS),HIWORD(pVer->dwProductVersionLS));
		}
		else
			Sprintf(title,_countof(title),L"Settings for Classic Start Menu");
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

		SendDlgItemMessage(hwndDlg,IDC_COMBOSCROLL,CB_ADDSTRING,0,(LPARAM)L"Don't scroll (use multiple columns)");
		SendDlgItemMessage(hwndDlg,IDC_COMBOSCROLL,CB_ADDSTRING,0,(LPARAM)L"Scroll (use single column)");
		SendDlgItemMessage(hwndDlg,IDC_COMBOSCROLL,CB_ADDSTRING,0,(LPARAM)L"Auto (multiple columns if they fit)");

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
		const wchar_t *text[]={L"Nothing",L"Classic Menu",L"Windows Menu"};
		for (int i=0;i<_countof(text);i++)
		{
			SendDlgItemMessage(hwndDlg,IDC_COMBOCLICK,CB_ADDSTRING,0,(LPARAM)text[i]);
			SendDlgItemMessage(hwndDlg,IDC_COMBOSCLICK,CB_ADDSTRING,0,(LPARAM)text[i]);
			SendDlgItemMessage(hwndDlg,IDC_COMBOWIN,CB_ADDSTRING,0,(LPARAM)text[i]);
			SendDlgItemMessage(hwndDlg,IDC_COMBOSWIN,CB_ADDSTRING,0,(LPARAM)text[i]);
			SendDlgItemMessage(hwndDlg,IDC_COMBOMCLICK,CB_ADDSTRING,0,(LPARAM)text[i]);
		}
		SendDlgItemMessage(hwndDlg,IDC_COMBOCLICK,CB_SETCURSEL,settings.Controls&15,0);
		SendDlgItemMessage(hwndDlg,IDC_COMBOSCLICK,CB_SETCURSEL,(settings.Controls>>8)&15,0);
		SendDlgItemMessage(hwndDlg,IDC_COMBOWIN,CB_SETCURSEL,(settings.Controls>>16)&15,0);
		SendDlgItemMessage(hwndDlg,IDC_COMBOSWIN,CB_SETCURSEL,(settings.Controls>>24)&15,0);
		SendDlgItemMessage(hwndDlg,IDC_COMBOMCLICK,CB_SETCURSEL,(settings.Controls>>4)&15,0);

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

		SendDlgItemMessage(hwndDlg,IDC_LISTOPTIONS,LVM_SETEXTENDEDLISTVIEWSTYLE,LVS_EX_CHECKBOXES,LVS_EX_CHECKBOXES);
		RECT rc;
		GetClientRect(GetDlgItem(hwndDlg,IDC_LISTOPTIONS),&rc);
		LVCOLUMN column={LVCF_FMT|LVCF_WIDTH,LVCFMT_LEFT,rc.right-GetSystemMetrics(SM_CXVSCROLL)};
		SendDlgItemMessage(hwndDlg,IDC_LISTOPTIONS,LVM_INSERTCOLUMN,0,(LPARAM)&column);

		POINT pt={0,0};
		GetWindowRect(GetDlgItem(hwndDlg,IDC_STATICOPT),&rc);
		pt.y=rc.top;
		ScreenToClient(hwndDlg,&pt);
		g_VariationY=pt.y;
		GetWindowRect(GetDlgItem(hwndDlg,IDC_STATICVAR),&rc);
		pt.y=rc.top;
		ScreenToClient(hwndDlg,&pt);
		g_NoVariationY=pt.y;

		int idx=(int)SendDlgItemMessage(hwndDlg,IDC_COMBOSKIN,CB_ADDSTRING,0,(LPARAM)L"<Default>");
		SendDlgItemMessage(hwndDlg,IDC_COMBOSKIN,CB_SETCURSEL,idx,0);
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
		return TRUE;
	}

	if (uMsg==WM_COMMAND && wParam==MAKEWPARAM(IDC_COMBOSKIN,CBN_SELENDOK))
	{
		InitSkinVariations(hwndDlg,NULL,std::vector<unsigned int>());
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
		Sprintf(caption,_countof(caption),L"About skin %s",name);
		if (!LoadMenuSkin(name,skin,NULL,std::vector<unsigned int>(),true))
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
			wchar_t path[_MAX_PATH];
			GetModuleFileName(g_Instance,path,_countof(path));
			*PathFindFileName(path)=0;
			Strcat(path,_countof(path),DOC_PATH L"ClassicStartMenu.html");
			ShellExecute(hwndDlg,NULL,path,NULL,NULL,SW_SHOWNORMAL);
			return TRUE;
		}
		if (pHdr->idFrom==IDC_LINKINI && (pHdr->code==NM_CLICK || pHdr->code==NM_RETURN))
		{
			CRegKey regSettings;
			if (regSettings.Open(HKEY_CURRENT_USER,L"Software\\IvoSoft\\ClassicStartMenu")!=ERROR_SUCCESS)
				regSettings.Create(HKEY_CURRENT_USER,L"Software\\IvoSoft\\ClassicStartMenu");

			DWORD show;
			if (regSettings.QueryDWORDValue(L"IgnoreIniWarning",show)!=ERROR_SUCCESS)
			{
				TASKDIALOGCONFIG task={sizeof(task),hwndDlg,NULL,TDF_ALLOW_DIALOG_CANCELLATION,TDCBF_OK_BUTTON};
				task.pszMainIcon=TD_INFORMATION_ICON;
				task.pszWindowTitle=L"Classic Start Menu";
				task.pszMainInstruction=L"After modifying the ini file you have to save it and restart the Classic Start Menu. Right-click on the start button and select \"Exit\". Then run ClassicStartMenu.exe again. It will read the new settings.\n\nRemember: All lines starting with a semicolon are ignored. Remove the semicolon from the settings you want to use.";
				task.pszVerificationText=L"Don't show this message again";
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
