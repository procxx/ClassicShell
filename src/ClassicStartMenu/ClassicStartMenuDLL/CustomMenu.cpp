// Classic Shell (c) 2009-2011, Ivo Beltchev
// The sources for Classic Shell are distributed under the MIT open source license

#include "stdafx.h"
#include "CustomMenu.h"
#include "SettingsParser.h"
#include "Translations.h"
#include "MenuContainer.h"
#include "Settings.h"
#include "FNVHash.h"
#include "ResourceHelper.h"

// This table defines the standard menu items
static StdMenuItem g_StdMenu[]=
{
	// * means the command is not executable (for things like Settings, or for items that have FOLDERID)
	{L"*programs",MENU_PROGRAMS,&FOLDERID_Programs,&FOLDERID_CommonPrograms},
	{L"*favorites",MENU_FAVORITES,&FOLDERID_Favorites},
	{L"*documents",MENU_DOCUMENTS,&FOLDERID_Recent},
	{L"*settings",MENU_SETTINGS},
	{L"*search",MENU_SEARCH},
	{L"help",MENU_HELP},
	{L"run",MENU_RUN},
	{L"logoff",MENU_LOGOFF},
	{L"undock",MENU_UNDOCK},
	{L"disconnect",MENU_DISCONNECT},
	{L"shutdown_box",MENU_SHUTDOWN_BOX},
	{L"*user_files",MENU_USERFILES,&FOLDERID_UsersFiles},
	{L"*user_documents",MENU_USERDOCUMENTS,&FOLDERID_Documents},
	{L"*user_pictures",MENU_USERPICTURES,&FOLDERID_Pictures},
	{L"*control_panel",MENU_CONTROLPANEL,&FOLDERID_ControlPanelFolder},
	{L"*control_panel_categories",MENU_CONTROLPANEL_CATEGORIES,&FOLDERID_ControlPanelFolder},
	{L"windows_security",MENU_SECURITY},
	{L"*network_connections",MENU_NETWORK,&FOLDERID_ConnectionsFolder},
	{L"*printers",MENU_PRINTERS,&FOLDERID_PrintersFolder},
	{L"taskbar_settings",MENU_TASKBAR},
	{L"programs_features",MENU_FEATURES},
	{L"menu_settings",MENU_CLASSIC_SETTINGS},
	{L"search_files",MENU_SEARCH_FILES},
	{L"search_printer",MENU_SEARCH_PRINTER},
	{L"search_computers",MENU_SEARCH_COMPUTERS},
	{L"search_people",MENU_SEARCH_PEOPLE},
	{L"sleep",MENU_SLEEP},
	{L"hibernate",MENU_HIBERNATE},
	{L"restart",MENU_RESTART},
	{L"shutdown",MENU_SHUTDOWN},
	{L"switch_user",MENU_SWITCHUSER},
	{L"lock",MENU_LOCK},
	{L"*recent_items",MENU_RECENT_ITEMS},
	{L"search_box",MENU_SEARCH_BOX},
};

// {52528a6b-b9e3-4add-b60d-58 8c 2d ba 84 2d} (define homegroup GUID, so we don't need the Windows 7 SDK to compile Classic Shell)
static const GUID FOLDERID_HomeGroup2={0x52528a6b, 0xb9e3, 0x4add, {0xb6, 0x0d, 0x58, 0x8c, 0x2d, 0xba, 0x84, 0x2d}};

// This table defines folders that need special treatment
SpecialFolder g_SpecialFolders[]=
{
	{&FOLDERID_Games,SpecialFolder::FOLDER_NONEWFOLDER},
	{&FOLDERID_ComputerFolder,SpecialFolder::FOLDER_NONEWFOLDER},
	{&FOLDERID_RecycleBinFolder,SpecialFolder::FOLDER_NOSUBFOLDERS|SpecialFolder::FOLDER_NODROP},
	{&FOLDERID_NetworkFolder,SpecialFolder::FOLDER_NODROP},
	{&FOLDERID_ConnectionsFolder,SpecialFolder::FOLDER_NODROP},
	{&FOLDERID_Recent,SpecialFolder::FOLDER_NODROP},
	{&FOLDERID_ControlPanelFolder,SpecialFolder::FOLDER_NODROP},
	{&FOLDERID_PrintersFolder,SpecialFolder::FOLDER_NODROP},
	{&FOLDERID_HomeGroup2,SpecialFolder::FOLDER_NODROP},
	{NULL}
};

static std::vector<StdMenuItem> g_CustomMenu;
static unsigned int g_RootSettings;
static unsigned int g_MenuItemsHash;
static CSettingsParser g_CustomMenuParser;

static const StdMenuItem *FindStdMenuItem( const wchar_t *command )
{
	for (int i=0;i<_countof(g_StdMenu);i++)
	{
		const wchar_t *cmd=g_StdMenu[i].command;
		if (*cmd=='*') cmd++;
		if (_wcsicmp(cmd,command)==0)
			return &g_StdMenu[i];
	}
	return NULL;
}

static unsigned int ParseItemSettings( const wchar_t *name )
{
	wchar_t buf[256];
	Sprintf(buf,_countof(buf),L"%s.Settings",name);
	const wchar_t *str=g_CustomMenuParser.FindSetting(buf);
	if (!str) return 0;

	unsigned int settings=0;
	while(*str)
	{
		wchar_t token[256];
		str=GetToken(str,token,_countof(token),L", \t|;");
		if (_wcsicmp(token,L"OPEN_UP")==0) settings|=StdMenuItem::MENU_OPENUP;
		if (_wcsicmp(token,L"OPEN_UP_CHILDREN")==0) settings|=StdMenuItem::MENU_OPENUP_REC;
		if (_wcsicmp(token,L"SORT_ZA")==0) settings|=StdMenuItem::MENU_SORTZA;
		if (_wcsicmp(token,L"SORT_ZA_CHILDREN")==0) settings|=StdMenuItem::MENU_SORTZA_REC;
		if (_wcsicmp(token,L"SORT_ONCE")==0) settings|=StdMenuItem::MENU_SORTONCE;
		if (_wcsicmp(token,L"ITEMS_FIRST")==0) settings|=StdMenuItem::MENU_ITEMS_FIRST;
		if (_wcsicmp(token,L"TRACK_RECENT")==0) settings|=StdMenuItem::MENU_TRACK;
		if (_wcsicmp(token,L"NOTRACK_RECENT")==0) settings|=StdMenuItem::MENU_NOTRACK;
		if (_wcsicmp(token,L"NOEXPAND")==0) settings|=StdMenuItem::MENU_NOEXPAND;
		if (_wcsicmp(token,L"MULTICOLUMN")==0) settings|=StdMenuItem::MENU_MULTICOLUMN;
		if (_wcsicmp(token,L"NOEXTENSIONS")==0) settings|=StdMenuItem::MENU_NOEXTENSIONS;
		if (_wcsicmp(token,L"INLINE")==0) settings|=StdMenuItem::MENU_INLINE;
	}
	return settings;
}

static void ParseMenuItem( StdMenuItem &item, const wchar_t *name )
{
	wchar_t buf[1024];
	const wchar_t *str;
	Sprintf(buf,_countof(buf),L"%s.Link",name);
	str=g_CustomMenuParser.FindSetting(buf);
	if (str)
	{
		// parse link
		item.link=str;
		const wchar_t *c=wcschr(item.link,'|');
		if (c)
		{
			for (c++;*c==' ';)
				c++;
			item.link=c;
		}
	}

	Sprintf(buf,_countof(buf),L"%s.Command",name);
	str=g_CustomMenuParser.FindSetting(buf);
	if (str)
	{
		// parse command
		const StdMenuItem *pItem=FindStdMenuItem(str);
		if (pItem)
		{
			item.id=pItem->id;
			item.folder1=pItem->folder1;
			item.folder2=pItem->folder2;
			if (item.id==MENU_CONTROLPANEL_CATEGORIES)
			{
				item.id=MENU_CONTROLPANEL;
				item.command=L"::{26EE0668-A00A-44D7-9371-BEB064C98683}";
			}
			else if (*pItem->command!='*')
				item.command=pItem->command;
		}
		else
		{
			item.id=MENU_CUSTOM;
			item.command=str;
		}
	}

	Sprintf(buf,_countof(buf),L"%s.Label",name);
	str=g_CustomMenuParser.FindSetting(buf);
	if (str)
	{
		// parse name
		if (*str=='$')
		{
			item.label=FindTranslation(str+1,NULL);
			if (!item.label)
				item.label=str;
		}
		else
			item.label=str;
	}

	Sprintf(buf,_countof(buf),L"%s.Tip",name);
	str=g_CustomMenuParser.FindSetting(buf);
	if (str)
	{
		// parse name
		if (*str=='$')
			item.tip=FindTranslation(str+1,NULL);
		else
			item.tip=str;
	}

	Sprintf(buf,_countof(buf),L"%s.Icon",name);
	item.iconPath=g_CustomMenuParser.FindSetting(buf);

	item.settings=ParseItemSettings(name);
}

const StdMenuItem *ParseCustomMenu( unsigned int &rootSettings )
{
	CString menuText=GetSettingString(L"MenuItems");
	unsigned int hash=CalcFNVHash(menuText);
	if (hash!=g_MenuItemsHash)
	{
		g_RootSettings=0;
		g_MenuItemsHash=hash;
		g_CustomMenu.clear();
		g_CustomMenuParser.Reset();
		g_CustomMenuParser.LoadText(menuText,menuText.GetLength());
		g_CustomMenuParser.ParseText();

		std::vector<CSettingsParser::TreeItem> items;
		g_CustomMenuParser.ParseTree(L"Items",items);
		g_CustomMenu.resize(items.size());
		for (size_t i=0;i<items.size();i++)
		{
			const wchar_t *name=items[i].name;
			StdMenuItem &item=g_CustomMenu[i];

			item.command=0;
			item.id=MENU_NO;
			item.folder1=item.folder2=NULL;
			item.label=item.tip=item.iconPath=item.link=NULL;
			item.settings=0;
			item.submenu=NULL;

			// handle special names
			if (!*name)
			{
				item.id=MENU_LAST;
				continue;
			}
			if (_wcsicmp(name,L"SEPARATOR")==0)
			{
				item.id=MENU_SEPARATOR;
				continue;
			}
			if (_wcsicmp(name,L"COLUMN_PADDING")==0)
			{
				item.id=MENU_COLUMN_PADDING;
				continue;
			}
			if (_wcsicmp(name,L"COLUMN_BREAK")==0)
			{
				item.id=MENU_COLUMN_BREAK;
				continue;
			}

			// handle custom items
			item.id=MENU_CUSTOM;
			ParseMenuItem(item,name);
			if (item.id==MENU_RECENT_ITEMS)
				g_RootSettings|=StdMenuItem::MENU_NORECENT;
			int idx=items[i].children;
			if (idx>=0)
				item.submenu=&g_CustomMenu[idx];
		}

		for (std::vector<StdMenuItem>::iterator it=g_CustomMenu.begin();it!=g_CustomMenu.end();++it)
			if (it->id==MENU_RECENT_ITEMS)
			{
				g_RootSettings|=StdMenuItem::MENU_NORECENT;
				break;
			}

		// if there is no break, add one after Programs
		if (!g_CustomMenu.empty())
		{
			bool bBreak=false;
			int after=-1;
			for (int i=0;g_CustomMenu[i].id!=MENU_LAST;i++)
			{
				if (g_CustomMenu[i].id==MENU_COLUMN_BREAK)
					bBreak=true;
				if (g_CustomMenu[i].id==MENU_PROGRAMS)
					after=i;
			}
			if (!bBreak && after>=0)
			{
				// add break
				StdMenuItem br={NULL,MENU_COLUMN_BREAK};
				const StdMenuItem *pBase=&g_CustomMenu[0];
				g_CustomMenu.insert(g_CustomMenu.begin()+after+1,br);

				// fix submenu pointers
				for (std::vector<StdMenuItem>::iterator it=g_CustomMenu.begin();it!=g_CustomMenu.end();++it)
					if (it->submenu)
					{
						int idx=(int)(it->submenu-pBase);
						if (idx>after+1)
							idx++;
						it->submenu=&g_CustomMenu[idx];
					}
			}
		}
		// ignore extra search boxes
		bool bSearchBox=false;
		for (std::vector<StdMenuItem>::iterator it=g_CustomMenu.begin();it!=g_CustomMenu.end();++it)
		{
			if (it->id==MENU_SEARCH_BOX)
			{
				if (!bSearchBox)
					bSearchBox=true;
				else
					it->id=MENU_IGNORE;
			}
		}
	}

	rootSettings=g_RootSettings;
	return &g_CustomMenu[0];
}
