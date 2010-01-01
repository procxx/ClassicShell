// Classic Shell (c) 2009-2010, Ivo Beltchev
// The sources for Classic Shell are distributed under the MIT open source license

#include "stdafx.h"
#include "ClassicStartMenuDLL.h"
#include "MenuContainer.h"
#include "IconManager.h"
#include "ParseSettings.h"
#include "Settings.h"

#define HOOK_DROPTARGET // define this to replace the IDropTarget of the start button

#if defined(BUILD_SETUP) && !defined(HOOK_DROPTARGET)
#define HOOK_DROPTARGET // make sure it is defined in Setup
#endif

HWND g_StartButton;
HWND g_TaskBar;
static UINT g_StartMenuMsg;
static HWND g_Tooltip;
static TOOLINFO g_StarButtonTool;
static int g_HotkeyID;
static DWORD g_Hotkey;

static BOOL CALLBACK FindTooltipEnum( HWND hwnd, LPARAM lParam )
{
	// look for tooltip control in the current thread that has a tool for g_TaskBar+g_StartButton
	wchar_t name[256];
	GetClassName(hwnd,name,_countof(name));
	if (_wcsicmp(name,TOOLTIPS_CLASS)!=0) return TRUE;
	TOOLINFO info={sizeof(info),0,g_TaskBar,(UINT_PTR)g_StartButton};
	if (SendMessage(hwnd,TTM_GETTOOLINFO,0,(LPARAM)&info))
	{
		g_Tooltip=hwnd;
		return FALSE;
	}
	return TRUE;
}

static BOOL CALLBACK FindStartButtonEnum( HWND hwnd, LPARAM lParam )
{
	// look for top-level window in the current thread with class "button"
	wchar_t name[256];
	GetClassName(hwnd,name,_countof(name));
	if (_wcsicmp(name,L"button")!=0) return TRUE;
	g_StartButton=hwnd;
	return FALSE;
}

static BOOL CALLBACK FindTaskBarEnum( HWND hwnd, LPARAM lParam )
{
	// look for top-level window with class "Shell_TrayWnd" and process ID=lParam
	DWORD process;
	GetWindowThreadProcessId(hwnd,&process);
	if (process!=lParam) return TRUE;
	wchar_t name[256];
	GetClassName(hwnd,name,_countof(name));
	if (_wcsicmp(name,L"Shell_TrayWnd")!=0) return TRUE;
	g_TaskBar=hwnd;
	return FALSE;
}

// Find the start button window for the given process
STARTMENUAPI HWND FindStartButton( DWORD process )
{
	g_StartButton=NULL;
	g_TaskBar=NULL;
	g_Tooltip=NULL;
	// find the taskbar
	EnumWindows(FindTaskBarEnum,process);
	if (g_TaskBar)
	{
		// find start button
		EnumThreadWindows(GetWindowThreadProcessId(g_TaskBar,NULL),FindStartButtonEnum,NULL);
		// find tooltip
		if (GetWindowThreadProcessId(g_TaskBar,NULL)==GetCurrentThreadId())
			EnumThreadWindows(GetWindowThreadProcessId(g_TaskBar,NULL),FindTooltipEnum,NULL);
		if (g_Tooltip)
		{
			g_StarButtonTool.cbSize=sizeof(g_StarButtonTool);
			g_StarButtonTool.hwnd=g_TaskBar;
			g_StarButtonTool.uId=(UINT_PTR)g_StartButton;
			SendMessage(g_Tooltip,TTM_GETTOOLINFO,0,(LPARAM)&g_StarButtonTool);
		}
	}
	return g_StartButton;
}

#ifdef HOOK_DROPTARGET
class CStartMenuTarget: public IDropTarget
{
public:
	CStartMenuTarget( void ) { m_RefCount=1; }
	// IUnknown
	virtual STDMETHODIMP QueryInterface( REFIID riid, void **ppvObject )
	{
		*ppvObject=NULL;
		if (IID_IUnknown==riid || IID_IDropTarget==riid)
		{
			AddRef();
			*ppvObject=(IDropTarget*)this;
			return S_OK;
		}
		return E_NOINTERFACE;
	}

	virtual ULONG STDMETHODCALLTYPE AddRef( void ) 
	{ 
		return InterlockedIncrement(&m_RefCount);
	}

	virtual ULONG STDMETHODCALLTYPE Release( void )
	{
		long nTemp=InterlockedDecrement(&m_RefCount);
		if (!nTemp) delete this;
		return nTemp;
	}

	// IDropTarget
	virtual HRESULT STDMETHODCALLTYPE DragEnter( IDataObject *pDataObj, DWORD grfKeyState, POINTL pt, DWORD *pdwEffect )
	{
		PostMessage(g_StartButton,g_StartMenuMsg,1,0);
		*pdwEffect=DROPEFFECT_NONE;
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE DragOver( DWORD grfKeyState, POINTL pt, DWORD *pdwEffect ) { return *pdwEffect=DROPEFFECT_NONE; return S_OK; }
	virtual HRESULT STDMETHODCALLTYPE DragLeave( void ) { return S_OK; }
	virtual HRESULT STDMETHODCALLTYPE Drop( IDataObject *pDataObj, DWORD grfKeyState, POINTL pt, DWORD *pdwEffect ) { return *pdwEffect=DROPEFFECT_NONE; return S_OK; }

private:
	LONG m_RefCount;
};

#endif

static CComPtr<IDropTarget> g_pOriginalTarget;

static void FindStartButton( void )
{
	if (!g_StartButton)
	{
		g_StartMenuMsg=RegisterWindowMessage(L"ClassicStartMenu.StartMenuMsg");
		g_StartButton=FindStartButton(GetCurrentProcessId());
#ifdef HOOK_DROPTARGET
		if (g_StartButton)
		{
			g_pOriginalTarget=(IDropTarget*)GetProp(g_StartButton,L"OleDropTargetInterface");
			if (g_pOriginalTarget)
				RevokeDragDrop(g_StartButton);
			CStartMenuTarget *pNewTarget=new CStartMenuTarget();
			RegisterDragDrop(g_StartButton,pNewTarget);
			pNewTarget->Release();
			g_HotkeyID=GlobalAddAtom(L"ClassicStartMenu.Hotkey");
			StartMenuSettings settings;
			ReadSettings(settings);
			SetHotkey(settings.Hotkey);
		}
#endif
		if (!g_StartButton) g_StartButton=(HWND)1;
	}
}

void EnableStartTooltip( bool bEnable )
{
	if (g_Tooltip)
	{
		SendMessage(g_Tooltip,TTM_POP,0,0);
		if (bEnable)
			SendMessage(g_Tooltip,TTM_UPDATETIPTEXT,0,(LPARAM)&g_StarButtonTool);
		else
		{
			TOOLINFO info=g_StarButtonTool;
			info.lpszText=L"";
			SendMessage(g_Tooltip,TTM_UPDATETIPTEXT,0,(LPARAM)&info);
		}
	}
}

// Restore the original drop target
void UnhookDropTarget( void )
{
#ifdef HOOK_DROPTARGET
	if (g_pOriginalTarget)
	{
		RevokeDragDrop(g_StartButton);
		RegisterDragDrop(g_StartButton,g_pOriginalTarget);
		g_pOriginalTarget=NULL;
		SetHotkey(0);
	}
#endif
}

// Set the hotkey for the start menu (0 - Win key, 1 - no hotkey)
void SetHotkey( DWORD hotkey )
{
	if (g_Hotkey>1)
		UnregisterHotKey(g_StartButton,g_HotkeyID);
	g_Hotkey=hotkey;
	if (hotkey>1)
	{
		int mod=0;
		if (hotkey&(HOTKEYF_SHIFT<<8)) mod|=MOD_SHIFT;
		if (hotkey&(HOTKEYF_CONTROL<<8)) mod|=MOD_CONTROL;
		if (hotkey&(HOTKEYF_ALT<<8)) mod|=MOD_ALT;
		RegisterHotKey(g_StartButton,g_HotkeyID,mod,hotkey&255);
	}
}

// Toggle the start menu. bKeyboard - set to true to show the keyboard cues
STARTMENUAPI HWND ToggleStartMenu( HWND startButton, bool bKeyboard )
{
	return CMenuContainer::ToggleStartMenu(startButton,bKeyboard);
}

///////////////////////////////////////////////////////////////////////////////

// WH_GETMESSAGE hook for the Progman window
STARTMENUAPI LRESULT CALLBACK HookProgMan( int code, WPARAM wParam, LPARAM lParam )
{
	if (code==HC_ACTION)
	{
		MSG *msg=(MSG*)lParam;
		if (msg->message==WM_SYSCOMMAND && (msg->wParam&0xFFF0)==SC_TASKLIST)
		{
			FindStartButton();
			if (g_Hotkey==0)
			{
				if (g_StartButton)
				{
					SetForegroundWindow(g_StartButton);
					PostMessage(g_StartButton,g_StartMenuMsg,0,0);
				}
				msg->message=WM_NULL;
			}
		}
	}
	return CallNextHookEx(NULL,code,wParam,lParam);
}

static bool g_bInMenu=false;
static DWORD g_LastClickTime=0;

// WH_GETMESSAGE hook for the start button window
STARTMENUAPI LRESULT CALLBACK HookStartButton( int code, WPARAM wParam, LPARAM lParam )
{
	if (code==HC_ACTION && !g_bInMenu)
	{
		MSG *msg=(MSG*)lParam;
		FindStartButton();
		if (IsSettingsMessage(msg))
		{
			msg->message=WM_NULL;
			return 0;
		}
		if (msg->message==g_StartMenuMsg && msg->hwnd==g_StartButton)
		{
			// wParam=0 - toggle
			// wParam=1 - open
			msg->message=WM_NULL;
			if (msg->wParam==0 || !CMenuContainer::IsMenuOpened())
				ToggleStartMenu(g_StartButton,true);
		}
		if (msg->message==WM_HOTKEY && msg->hwnd==g_StartButton && msg->wParam==g_HotkeyID)
		{
			msg->message=WM_NULL;
			SetForegroundWindow(g_StartButton);
			ToggleStartMenu(g_StartButton,true);
		}

		if ((msg->message==WM_LBUTTONDOWN || msg->message==WM_LBUTTONDBLCLK) && msg->hwnd==g_StartButton && !(msg->wParam&MK_SHIFT))
		{
			DWORD pos=GetMessagePos();
			POINT pt={(short)LOWORD(pos),(short)HIWORD(pos)};
			if (msg->time==g_LastClickTime || WindowFromPoint(pt)!=g_StartButton)
			{
				// ignore the click if it matches the last click's timestamp (sometimes the same message comes twice)
				// or when the mouse is not over the start button (sometimes clicks on a context menu are sent to the start button)
				return CallNextHookEx(NULL,code,wParam,lParam);
			}
			g_LastClickTime=msg->time;
			// click on the start button - toggle the menu
			DWORD keyboard;
			SystemParametersInfo(SPI_GETKEYBOARDCUES,NULL,&keyboard,0);
			ToggleStartMenu(g_StartButton,keyboard!=0);
			msg->message=WM_NULL;
		}

		if ((msg->message==WM_NCLBUTTONDOWN || msg->message==WM_NCLBUTTONDBLCLK) && msg->wParam==HTCAPTION && msg->hwnd==g_TaskBar && GetKeyState(VK_SHIFT)>=0)
		{
			DWORD pos=GetMessagePos();
			POINT pt={(short)LOWORD(pos),(short)HIWORD(pos)};
			ScreenToClient(g_TaskBar,&pt);
			HWND child=ChildWindowFromPoint(g_TaskBar,pt);
			if (child!=NULL && child!=g_TaskBar)
			{
				// ignore the click if it is on a child window (like the rebar or the tray area)
				return CallNextHookEx(NULL,code,wParam,lParam);
			}
			// click on the taskbar around the start menu - toggle the menu
			DWORD keyboard;
			SystemParametersInfo(SPI_GETKEYBOARDCUES,NULL,&keyboard,0);
			ToggleStartMenu(g_StartButton,keyboard!=0);
			msg->message=WM_NULL;
		}

		if (msg->message==WM_TIMER && msg->hwnd==g_TaskBar && CMenuContainer::IgnoreTaskbarTimers())
		{
			// stop the taskbar timer messages. prevents the auto-hide taskbar from closing
			msg->message=WM_NULL;
		}

		if (msg->message==WM_RBUTTONUP && msg->hwnd==g_StartButton && !(msg->wParam&MK_SHIFT))
		{
			// additional commands for the context menu
			enum
			{
				CMD_SETTINGS=1,
				CMD_HELP,
				CMD_EXIT,
				CMD_OPEN,
				CMD_OPEN_ALL,
			};

			// right-click on the start button - open the context menu (Settings, Help, Exit)
			msg->message=WM_NULL;
			POINT p={(short)LOWORD(msg->lParam),(short)HIWORD(msg->lParam)};
			ClientToScreen(msg->hwnd,&p);
			HMENU menu=CreatePopupMenu();
			AppendMenu(menu,MF_STRING,0,L"== Classic Start Menu ==");
			EnableMenuItem(menu,0,MF_BYPOSITION|MF_DISABLED);
			SetMenuDefaultItem(menu,0,TRUE);
			AppendMenu(menu,MF_SEPARATOR,0,0);
			AppendMenu(menu,MF_STRING,CMD_OPEN,FindSetting("Menu.Open",L"&Open"));
			AppendMenu(menu,MF_STRING,CMD_OPEN_ALL,FindSetting("Menu.OpenAll",L"O&pen All Users"));
			AppendMenu(menu,MF_SEPARATOR,0,0);
			AppendMenu(menu,MF_STRING,CMD_SETTINGS,FindSetting("Menu.MenuSettings",L"Settings"));
			AppendMenu(menu,MF_STRING,CMD_HELP,FindSetting("Menu.MenuHelp",L"Help"));
			AppendMenu(menu,MF_STRING,CMD_EXIT,FindSetting("Menu.MenuExit",L"Exit"));
			MENUITEMINFO mii={sizeof(mii)};
			mii.fMask=MIIM_BITMAP;
			mii.hbmpItem=HBMMENU_POPUP_CLOSE;
			SetMenuItemInfo(menu,3,FALSE,&mii);
			g_bInMenu=true;
			int res=TrackPopupMenu(menu,TPM_LEFTBUTTON|TPM_RETURNCMD|(IsLanguageRTL()?TPM_LAYOUTRTL:0),p.x,p.y,0,msg->hwnd,NULL);
			DestroyMenu(menu);
			g_bInMenu=false;
			if (res==CMD_SETTINGS)
			{
				EditSettings(false);
			}
			if (res==CMD_HELP)
			{
				wchar_t path[_MAX_PATH];
				GetModuleFileName(g_Instance,path,_countof(path));
				*PathFindFileName(path)=0;
				wcscat_s(path,L"ClassicStartMenu.html");
				ShellExecute(NULL,NULL,path,NULL,NULL,SW_SHOWNORMAL);
			}
			if (res==CMD_EXIT)
			{
				// cleanup
				CloseSettings();
				CMenuContainer::CloseStartMenu();
				g_IconManager.StopPreloading(true);
				UnhookDropTarget();
				// send WM_CLOSE to the window in ClassicStartMenu.exe. it will unhook everything and unload the DLL
				HWND hwnd=FindWindow(L"ClassicStartMenu.CStartHookWindow",L"StartHookWindow");
				if (hwnd) PostMessage(hwnd,WM_CLOSE,0,0);
			}
			if (res==CMD_OPEN || res==CMD_OPEN_ALL)
			{
				wchar_t *path;
				if (SUCCEEDED(SHGetKnownFolderPath((res==CMD_OPEN)?FOLDERID_StartMenu:FOLDERID_CommonStartMenu,0,NULL,&path)))
				{
					ShellExecute(NULL,L"open",path,NULL,NULL,SW_SHOWNORMAL);
					CoTaskMemFree(path);
				}
			}
		}

	}
	return CallNextHookEx(NULL,code,wParam,lParam);
}
