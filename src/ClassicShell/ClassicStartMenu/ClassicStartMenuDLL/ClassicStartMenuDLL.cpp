// Classic Shell (c) 2009, Ivo Beltchev
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
	}
#endif
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
			if (g_StartButton)
			{
				SetForegroundWindow(g_StartButton);
				PostMessage(g_StartButton,g_StartMenuMsg,0,0);
			}
			msg->message=WM_NULL;
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
		if (msg->message==g_StartMenuMsg && msg->hwnd==g_StartButton)
		{
			// wParam=0 - toggle
			// wParam=1 - open
			msg->message=WM_NULL;
			if (msg->wParam==0 || !CMenuContainer::IsMenuOpened())
				ToggleStartMenu(g_StartButton,true);
		}

		if ((msg->message==WM_LBUTTONDOWN || msg->message==WM_LBUTTONDBLCLK) && msg->hwnd==g_StartButton && GetKeyState(VK_SHIFT)>=0)
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

		if (msg->message==WM_RBUTTONUP && msg->hwnd==g_StartButton && GetKeyState(VK_SHIFT)>=0)
		{
			// right-click on the start button - open the context menu (Settings, Help, Exit)
			msg->message=WM_NULL;
			POINT p={(short)LOWORD(msg->lParam),(short)HIWORD(msg->lParam)};
			ClientToScreen(msg->hwnd,&p);
			HMENU menu=CreatePopupMenu();
			AppendMenu(menu,MF_STRING,0,L"== Classic Start Menu ==");
			EnableMenuItem(menu,0,MF_BYPOSITION|MF_DISABLED);
			SetMenuDefaultItem(menu,0,TRUE);
			AppendMenu(menu,MF_SEPARATOR,0,0);
			AppendMenu(menu,MF_STRING,1,FindSetting("Menu.MenuSettings",L"Settings"));
			AppendMenu(menu,MF_STRING,2,FindSetting("Menu.MenuHelp",L"Help"));
			AppendMenu(menu,MF_STRING,3,FindSetting("Menu.MenuExit",L"Exit"));
			g_bInMenu=true;
			int res=TrackPopupMenu(menu,TPM_LEFTBUTTON|TPM_RETURNCMD|(IsLanguageRTL()?TPM_LAYOUTRTL:0),p.x,p.y,0,msg->hwnd,NULL);
			DestroyMenu(menu);
			g_bInMenu=false;
			if (res==1) // settings
			{
				EditSettings();
			}
			if (res==2) // help
			{
				wchar_t path[_MAX_PATH];
				GetModuleFileName(g_Instance,path,_countof(path));
				*PathFindFileName(path)=0;
				wcscat_s(path,L"ClassicStartMenu.html");
				ShellExecute(NULL,NULL,path,NULL,NULL,SW_SHOWNORMAL);
			}
			if (res==3) // exit
			{
				// cleanup
				CloseSettings();
				CMenuContainer::CloseStartMenu();
				g_IconManager.StopPreloading(true);
				UnhookDropTarget();
				// send WM_CLOSE to the window in ClassicStartMenu.exe. it will unhook everything and unload the DLL
				HWND hwnd=FindWindow(L"ClassicStartMenu.CStartHookWindow",L"StartHookWindow");
				if (hwnd) SendMessage(hwnd,WM_CLOSE,0,0);
			}
		}
	}
	return CallNextHookEx(NULL,code,wParam,lParam);
}
