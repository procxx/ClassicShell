// Classic Shell (c) 2009-2010, Ivo Beltchev
// The sources for Classic Shell are distributed under the MIT open source license

#include "stdafx.h"
#include "CustomMenu.h"
#include "ParseSettings.h"
#include "TranslationSettings.h"
#include "ParseSettings.h"
#include "MenuContainer.h"

// This table defines the standard menu items
static StdMenuItem g_StdMenu[]=
{
	// Start menu
	{MENU_COLUMN_PADDING},
	{MENU_PROGRAMS,"Menu.Programs",L"&Programs",326,MENU_NO,&FOLDERID_Programs,&FOLDERID_CommonPrograms},
	{MENU_COLUMN_BREAK},
	{MENU_FAVORITES,"Menu.Favorites",L"F&avorites",322,MENU_NO,&FOLDERID_Favorites},
	{MENU_DOCUMENTS,"Menu.Documents",L"&Documents",327,MENU_USERFILES,&FOLDERID_Recent},
	{MENU_SETTINGS,"Menu.Settings",L"&Settings",330,MENU_CONTROLPANEL},
	{MENU_SEARCH,"Menu.Search",L"Sear&ch",323,MENU_SEARCH_FILES},
	{MENU_HELP,"Menu.Help",L"&Help and Support",324},
	{MENU_RUN,"Menu.Run",L"&Run...",328},
	{MENU_COLUMN_PADDING},
	{MENU_SEPARATOR},
	{MENU_LOGOFF,"Menu.Logoff",L"&Log Off %s...",325},
	{MENU_UNDOCK,"Menu.Undock",L"Undock Comput&er",331},
	{MENU_DISCONNECT,"Menu.Disconnect",L"D&isconnect...",329},
	{MENU_SHUTDOWN_BOX,"Menu.Shutdown",L"Sh&ut Down...",329},
	{MENU_LAST},

	// Documents
	{MENU_USERFILES,NULL,NULL,0,MENU_NO,&FOLDERID_UsersFiles,NULL,"Menu.UserFilesTip",L"Contains folders for Documents, Pictures, Music, and other files that belong to you."},
	{MENU_USERDOCUMENTS,NULL,NULL,0,MENU_NO,&FOLDERID_Documents,NULL,"Menu.UserDocumentsTip",L"Contains letters, reports, and other documents and files."},
	{MENU_USERPICTURES,NULL,NULL,0,MENU_NO,&FOLDERID_Pictures,NULL,"Menu.UserPicturesTip",L"Contains digital photos, images, and graphic files."},
	{MENU_LAST},

	// Settings
	{MENU_CONTROLPANEL,"Menu.ControlPanel",L"&Control Panel",137,MENU_NO,&FOLDERID_ControlPanelFolder},
	{MENU_SEPARATOR},
	{MENU_SECURITY,"Menu.Security",L"Windows Security",48},
	{MENU_NETWORK,"Menu.Network",L"&Network Connections",257,MENU_NO,&FOLDERID_ConnectionsFolder,NULL,"Menu.NetworkTip",L"Displays existing network connections on this computer and helps you create new ones"},
	{MENU_PRINTERS,"Menu.Printers",L"&Printers",138,MENU_NO,&FOLDERID_PrintersFolder,NULL,"Menu.PrintersTip",L"Add, remove, and configure local and network printers."},
	{MENU_TASKBAR,"Menu.Taskbar",L"&Taskbar and Start Menu",40,MENU_NO,NULL,NULL,"Menu.TaskbarTip",L"Customize the Start Menu and the taskbar, such as the types of items to be displayed and how they should appear."},
	{MENU_FEATURES,"Menu.Features",L"Programs and &Features",271,MENU_NO,NULL,NULL,"Menu.FeaturesTip",L"Uninstall or change programs on your computer."},
	{MENU_SEPARATOR},
	{MENU_CLASSIC_SETTINGS,"Menu.ClassicSettings",L"Classic Start &Menu",274,MENU_NO,NULL,NULL,"Menu.SettingsTip",L"Settings for Classic Start Menu",NULL,NULL,NULL,L"ClassicStartMenuDLL.dll,103"},
	{MENU_LAST},

	// Search
	{MENU_SEARCH_FILES,"Menu.SearchFiles",L"For &Files and Folders...",134},
	{MENU_SEARCH_PRINTER,"Menu.SearchPrinter",L"For &Printer",1006},
	{MENU_SEARCH_COMPUTERS,"Menu.SearchComputers",L"For &Computers",135},
	{MENU_SEARCH_PEOPLE,"Menu.SearchPeople",L"For &People...",269},
	{MENU_LAST},
};

static std::vector<StdMenuItem> g_CustomMenu;
static unsigned int g_RootSettings;
static int g_CustomMenuRoot;
static FILETIME g_IniTimestamp;
static CSettingsParser g_CustomMenuParser;

static const StdMenuItem *FindStdMenuItem( TMenuID id )
{
	if (id!=MENU_NO && id!=MENU_SEPARATOR && id!=MENU_EMPTY && id!=MENU_EMPTY_TOP)
	{
		for (int i=0;i<_countof(g_StdMenu);i++)
			if (g_StdMenu[i].id==id)
				return &g_StdMenu[i];
	}
	return NULL;
}


static struct
{
	const wchar_t *name;
	TMenuID id;
}

g_StdItems[]={
	{L"PROGRAMS",MENU_PROGRAMS},
	{L"FAVORITES",MENU_FAVORITES},
	{L"DOCUMENTS",MENU_DOCUMENTS},
	{L"USER_FILES",MENU_USERFILES},
	{L"USER_DOCUMENTS",MENU_USERDOCUMENTS},
	{L"USER_PICTURES",MENU_USERPICTURES},
	{L"CONTROL_PANEL",MENU_CONTROLPANEL},
	{L"SECURITY",MENU_SECURITY},
	{L"NETWORK",MENU_NETWORK},
	{L"PRINTERS",MENU_PRINTERS},
},

g_StdCommands[]={
	{L"run",MENU_RUN},
	{L"help",MENU_HELP},
	{L"logoff",MENU_LOGOFF},
	{L"sleep",MENU_SLEEP},
	{L"hibernate",MENU_HIBERNATE},
	{L"restart",MENU_RESTART},
	{L"shutdown",MENU_SHUTDOWN},
	{L"switch_user",MENU_SWITCHUSER},
	{L"undock",MENU_UNDOCK},
	{L"disconnect",MENU_DISCONNECT},
	{L"shutdown_box",MENU_SHUTDOWN_BOX},
	{L"security",MENU_SECURITY},
	{L"search_files",MENU_SEARCH_FILES},
	{L"search_printer",MENU_SEARCH_PRINTER},
	{L"search_computers",MENU_SEARCH_COMPUTERS},
	{L"search_people",MENU_SEARCH_PEOPLE},
	{L"taskbar_settings",MENU_TASKBAR},
	{L"programs_settings",MENU_FEATURES},
	{L"menu_settings",MENU_CLASSIC_SETTINGS},
	{L"recent_items",MENU_RECENT_ITEMS},
};

static TMenuID FindStdItem( const wchar_t *name )
{
	for (int i=0;i<_countof(g_StdItems);i++)
		if (_wcsicmp(g_StdItems[i].name,name)==0)
			return g_StdItems[i].id;
	return MENU_NO;
}

static TMenuID FindStdCommand( const wchar_t *name )
{
	for (int i=0;i<_countof(g_StdCommands);i++)
		if (_wcsicmp(g_StdCommands[i].name,name)==0)
			return g_StdCommands[i].id;
	return MENU_NO;
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
	}
	return settings;
}

static void ParseMenuItem( StdMenuItem &item, const wchar_t *name )
{
	const StdMenuItem *pItem=FindStdMenuItem(FindStdItem(name));
	if (pItem) item=*pItem;

	wchar_t buf[1024];
	const wchar_t *str;
	Sprintf(buf,_countof(buf),L"%s.Link",name);
	str=g_CustomMenuParser.FindSetting(buf);
	if (str)
	{
		// parse link
		item.link=str;
	}

	Sprintf(buf,_countof(buf),L"%s.Command",name);
	str=g_CustomMenuParser.FindSetting(buf);
	if (str)
	{
		// parse command
		TMenuID id=FindStdCommand(str);
		if (id==MENU_NO)
		{
			if (item.id==MENU_NO)
				item.id=MENU_CUSTOM;
		}
		else
		{
			item.id=id;
			const StdMenuItem *pItem=FindStdMenuItem(id);
			if (pItem)
				item=*pItem;
		}
		item.command=str;
	}

	Sprintf(buf,_countof(buf),L"%s.Name",name);
	str=g_CustomMenuParser.FindSetting(buf);
	if (str)
	{
		// parse name
		item.key=NULL;
		if (*str=='$')
			item.name=FindTranslation(str+1,NULL);
		else
			item.name=str;
	}

	Sprintf(buf,_countof(buf),L"%s.Tip",name);
	str=g_CustomMenuParser.FindSetting(buf);
	if (str)
	{
		// parse name
		item.tipKey=NULL;
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
	static bool bInit=true;
	if (bInit)
	{
		bInit=false;
		for (int i=0;i<_countof(g_StdMenu);i++)
			g_StdMenu[i].submenu=FindStdMenuItem(g_StdMenu[i].submenuID);
	}

	wchar_t fname[_MAX_PATH];
	GetModuleFileName(g_Instance,fname,_countof(fname));
	*PathFindFileName(fname)=0;
	Strcat(fname,_countof(fname),INI_PATH L"StartMenuItems.ini");
	WIN32_FILE_ATTRIBUTE_DATA data;
	if (GetFileAttributesEx(fname,GetFileExInfoStandard,&data))
	{
		if (CompareFileTime(&g_IniTimestamp,&data.ftLastWriteTime)!=0)
		{
			g_IniTimestamp=data.ftLastWriteTime;
			g_CustomMenu.clear();
			g_CustomMenuParser.Reset();
			g_RootSettings=0;
			if (g_CustomMenuParser.LoadText(fname))
			{
				g_CustomMenuParser.ParseText();
				g_RootSettings=ParseItemSettings(L"MAIN_MENU");

				const wchar_t *str=g_CustomMenuParser.FindSetting("EnableCustomMenu");
				if (str && _wtol(str))
				{
					std::vector<CSettingsParser::TreeItem> items;
					g_CustomMenuParser.ParseTree(L"MAIN_MENU.Items",items);
					g_CustomMenu.resize(items.size());
					for (size_t i=0;i<items.size();i++)
					{
						const wchar_t *name=items[i].name;
						StdMenuItem &item=g_CustomMenu[i];
						memset(&item,0,sizeof(item));

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
							StdMenuItem br={MENU_COLUMN_BREAK};
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
				}
			}
		}
	}
	else
		g_CustomMenu.clear();

	if (g_CustomMenu.empty())
	{
		rootSettings=0;
		return g_StdMenu;
	}
	else
	{
		rootSettings=g_RootSettings;
		return &g_CustomMenu[0];
	}
}
