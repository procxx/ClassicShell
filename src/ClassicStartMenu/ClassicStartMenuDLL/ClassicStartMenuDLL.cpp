// Classic Shell (c) 2009-2012, Ivo Beltchev
// The sources for Classic Shell are distributed under the MIT open source license

#include "stdafx.h"
#include "resource.h"
#include "ClassicStartMenuDLL.h"
#include "ClassicStartButton.h"
#include "MenuContainer.h"
#include "IconManager.h"
#include "SettingsParser.h"
#include "Translations.h"
#include "Settings.h"
#include "SettingsUI.h"
#include "ResourceHelper.h"
#include "dllmain.h"
#include <uxtheme.h>
#include <htmlhelp.h>
#include <dbghelp.h>

#define HOOK_DROPTARGET // define this to replace the IDropTarget of the start button

#ifdef BUILD_SETUP
#ifndef HOOK_DROPTARGET
#define HOOK_DROPTARGET // make sure it is defined in Setup
#endif
#endif

HWND g_StartButton;
HWND g_TaskBar;
HWND g_OwnerWindow;
HWND g_ProgWin;
HWND g_TopMenu, g_AllPrograms, g_ProgramsButton, g_UserPic; // from the Windows menu
static bool g_bReplaceButton;
static HWND g_Rebar;
static SIZE g_RebarOffset;
static UINT g_StartMenuMsg;
static HWND g_Tooltip;
static TOOLINFO g_StartButtonTool;
static int g_HotkeyCSM, g_HotkeyWSM, g_HotkeyCSMID, g_HotkeyWSMID;
static HHOOK g_ProgHook, g_StartHook, g_KeyboardHook, g_AppManagerHook;
static bool g_bAllProgramsTimer;
static bool g_bStartButtonTimer;
static bool g_bInMenu;
static DWORD g_LastClickTime;
static DWORD g_LastHoverPos;
static bool g_bCrashDump;
static bool g_bStartScreen;
static SIZE HOT_CORNER_SIZE={15,15};
static int g_MetroModeCount;
static HWND g_StartButtonOld;
static DWORD g_StartButtonOldSizes[12];
const int FIRST_BUTTON_BITMAP=6801;
static HWND g_SwitchList;

enum
{
	OPEN_NOTHING,
	OPEN_CLASSIC,
	OPEN_WINDOWS
};

// MiniDumpNormal - minimal information
// MiniDumpWithDataSegs - include global variables
// MiniDumpWithFullMemory - include heap
MINIDUMP_TYPE MiniDumpType=MiniDumpNormal;

static DWORD WINAPI SaveCrashDump( void *pExceptionInfo )
{
	HMODULE dbghelp=NULL;
	{
		wchar_t path[_MAX_PATH]=L"%LOCALAPPDATA%";
		DoEnvironmentSubst(path,_countof(path));

		dbghelp=LoadLibrary(L"dbghelp.dll");

		LPCTSTR szResult = NULL;

		typedef BOOL (WINAPI *MINIDUMPWRITEDUMP)(HANDLE hProcess, DWORD dwPid, HANDLE hFile, MINIDUMP_TYPE DumpType,
			CONST PMINIDUMP_EXCEPTION_INFORMATION ExceptionParam,
			CONST PMINIDUMP_USER_STREAM_INFORMATION UserStreamParam,
			CONST PMINIDUMP_CALLBACK_INFORMATION CallbackParam
			);
		MINIDUMPWRITEDUMP dump=NULL;
		if (dbghelp)
			dump=(MINIDUMPWRITEDUMP)GetProcAddress(dbghelp,"MiniDumpWriteDump");
		if (dump)
		{
			HANDLE file;
			for (int i=1;;i++)
			{
				wchar_t fname[_MAX_PATH];
				Sprintf(fname,_countof(fname),L"%s\\CSM_Crash%d.dmp",path,i);
				file=CreateFile(fname,GENERIC_WRITE,0,NULL,CREATE_NEW,FILE_ATTRIBUTE_NORMAL,NULL);
				if (file!=INVALID_HANDLE_VALUE || GetLastError()!=ERROR_FILE_EXISTS) break;
			}
			if (file!=INVALID_HANDLE_VALUE)
			{
				_MINIDUMP_EXCEPTION_INFORMATION ExInfo;
				ExInfo.ThreadId = GetCurrentThreadId();
				ExInfo.ExceptionPointers = (_EXCEPTION_POINTERS*)pExceptionInfo;
				ExInfo.ClientPointers = NULL;

				dump(GetCurrentProcess(),GetCurrentProcessId(),file,MiniDumpType,&ExInfo,NULL,NULL);
				CloseHandle(file);
			}
		}
	}
	if (dbghelp) FreeLibrary(dbghelp);
	TerminateProcess(GetCurrentProcess(),10);
	return 0;
}

LONG _stdcall TopLevelFilter( _EXCEPTION_POINTERS *pExceptionInfo )
{
	if (pExceptionInfo->ExceptionRecord->ExceptionCode==EXCEPTION_STACK_OVERFLOW)
	{
		// start a new thread to get a fresh stack (hoping there is enough stack left for CreateThread)
		HANDLE thread=CreateThread(NULL,0,SaveCrashDump,pExceptionInfo,0,NULL);
		WaitForSingleObject(thread,INFINITE);
		CloseHandle(thread);
	}
	else
		SaveCrashDump(pExceptionInfo);
	return EXCEPTION_CONTINUE_SEARCH;
}

///////////////////////////////////////////////////////////////////////////////

// COwnerWindow - a special window used as owner for some UI elements, like the ones created by IContextMenu::InvokeCommand.
// A menu window cannot be used because it may disappear immediately after InvokeCommand. Some UI elements, like the UAC-related
// stuff can be created long after InvokeCommand returns and the menu may be deleted by then.
class COwnerWindow: public CWindowImpl<COwnerWindow>
{
public:
	DECLARE_WND_CLASS_EX(L"ClassicShell.CTempWindow",0,COLOR_MENU)

	// message handlers
	BEGIN_MSG_MAP( COwnerWindow )
		MESSAGE_HANDLER( WM_ACTIVATE, OnActivate )
	END_MSG_MAP()

protected:
	LRESULT OnActivate( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled )
	{
		if (LOWORD(wParam)!=WA_INACTIVE)
			return 0;

		if (CMenuContainer::s_pDragSource) return 0;

		// check if another menu window is being activated
		// if not, close all menus
		for (std::vector<CMenuContainer*>::const_iterator it=CMenuContainer::s_Menus.begin();it!=CMenuContainer::s_Menus.end();++it)
			if ((*it)->m_hWnd==(HWND)lParam)
				return 0;

		for (std::vector<CMenuContainer*>::reverse_iterator it=CMenuContainer::s_Menus.rbegin();it!=CMenuContainer::s_Menus.rend();++it)
			if (!(*it)->m_bDestroyed)
				(*it)->PostMessage(WM_CLOSE);

		return 0;
	}
};

static COwnerWindow g_Owner;

///////////////////////////////////////////////////////////////////////////////

STARTMENUAPI LRESULT CALLBACK HookProgMan( int code, WPARAM wParam, LPARAM lParam );
STARTMENUAPI LRESULT CALLBACK HookStartButton( int code, WPARAM wParam, LPARAM lParam );

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

// Find the taskbar window for the given process
STARTMENUAPI bool FindTaskBar( DWORD process )
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
		if (GetWindowThreadProcessId(g_TaskBar,NULL)==GetCurrentThreadId())
		{
			// find tooltip
			if (g_StartButton)
			{
				EnumThreadWindows(GetWindowThreadProcessId(g_TaskBar,NULL),FindTooltipEnum,NULL);
				if (g_Tooltip)
				{
					g_StartButtonTool.cbSize=sizeof(g_StartButtonTool);
					g_StartButtonTool.hwnd=g_TaskBar;
					g_StartButtonTool.uId=(UINT_PTR)g_StartButton;
					SendMessage(g_Tooltip,TTM_GETTOOLINFO,0,(LPARAM)&g_StartButtonTool);
				}
			}
			g_OwnerWindow=g_Owner.Create(NULL,0,0,WS_POPUP,WS_EX_TOOLWINDOW|WS_EX_TOPMOST);
		}
	}
	return g_TaskBar!=NULL;
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
		PostMessage(g_TaskBar,g_StartMenuMsg,(grfKeyState&MK_SHIFT)?MSG_SHIFTDRAG:MSG_DRAG,0);
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

static void FindTaskBar( void )
{
	if (!g_TaskBar)
	{
		g_StartMenuMsg=RegisterWindowMessage(L"ClassicStartMenu.StartMenuMsg");
		FindTaskBar(GetCurrentProcessId());
		if (g_TaskBar)
		{
			g_HotkeyCSMID=GlobalAddAtom(L"ClassicStartMenu.HotkeyCSM");
			g_HotkeyWSMID=GlobalAddAtom(L"ClassicStartMenu.HotkeyWSM");
			EnableHotkeys(HOTKEYS_NORMAL);
			srand(GetTickCount());
		}
		if (!g_TaskBar) g_TaskBar=(HWND)1;
	}
}

void EnableStartTooltip( bool bEnable )
{
	if (g_Tooltip)
	{
		SendMessage(g_Tooltip,TTM_POP,0,0);
		if (bEnable)
			SendMessage(g_Tooltip,TTM_UPDATETIPTEXT,0,(LPARAM)&g_StartButtonTool);
		else
		{
			TOOLINFO info=g_StartButtonTool;
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
		if (g_pOriginalTarget)
			RegisterDragDrop(g_StartButton,g_pOriginalTarget);
		g_pOriginalTarget=NULL;
	}
#endif
}

// Toggle the start menu. bKeyboard - set to true to show the keyboard cues
STARTMENUAPI HWND ToggleStartMenu( HWND startButton, bool bKeyboard )
{
	return CMenuContainer::ToggleStartMenu(startButton,bKeyboard,false);
}

// Returns true if the mouse is on the taskbar portion of the start button
bool PointAroundStartButton( void )
{
	CPoint pt(GetMessagePos());
	RECT rc;
	GetWindowRect(g_TaskBar,&rc);
	if (!PtInRect(&rc,pt))
		return false;

	APPBARDATA appbar={sizeof(appbar),g_TaskBar};
	SHAppBarMessage(ABM_GETTASKBARPOS,&appbar);

	if (g_bReplaceButton)
	{
		// if there is a custom start button, it determines the width of the area
		RECT rc2;
		GetWindowRect(g_StartButton,&rc2);
		if (appbar.uEdge==ABE_LEFT || appbar.uEdge==ABE_RIGHT)
		{
			rc2.left=rc.left;
			rc2.right=rc.right;
		}
		else
		{
			rc2.top=rc.top;
			rc2.bottom=rc.bottom;
		}
		return PtInRect(&rc2,pt)!=0;
	}
	else
	{
		// the Vista/Win7 start button

		// in Classic mode there are few pixels at the edge of the taskbar that are not covered by a child window
		// we nudge the point to be in the middle of the taskbar to avoid those pixels
		// also ignore clicks on the half of the taskbar that doesn't contain the start button
		if (appbar.uEdge==ABE_LEFT || appbar.uEdge==ABE_RIGHT)
		{
			pt.x=(rc.left+rc.right)/2; // vertical taskbar, set X
			if (pt.y>(rc.top+rc.bottom)/2)
				return false;
		}
		else
		{
			pt.y=(rc.top+rc.bottom)/2; // vertical taskbar, set Y
			if (pt.x>(rc.left+rc.right)/2)
			{
				if (!(GetWindowLong(g_TaskBar,GWL_EXSTYLE)&WS_EX_LAYOUTRTL))
					return false;
			}
			else
			{
				if (GetWindowLong(g_TaskBar,GWL_EXSTYLE)&WS_EX_LAYOUTRTL)
					return false;
			}
		}
		ScreenToClient(g_TaskBar,&pt);
		HWND child=ChildWindowFromPointEx(g_TaskBar,pt,CWP_SKIPINVISIBLE|CWP_SKIPTRANSPARENT);
		if (child!=NULL && child!=g_TaskBar)
		{
			// ignore the click if it is on a child window (like the rebar or the tray area)
			return false;
		}
		return true;
	}
}

#ifndef __IMetroMode_INTERFACE_DEFINED__
// declare the IMetroMode interface so we don't need the Win8 SDK
typedef enum METRO_MONITOR_MODE;
interface IMonitorModeEvents;

MIDL_INTERFACE("2246EA2D-CAEA-4444-A3C4-6DE827E44313")
IMetroMode: public IUnknown
{
public:
	virtual HRESULT STDMETHODCALLTYPE GetMonitorMode(
		/* [in] */ __RPC__in HMONITOR hMonitor,
		/* [out] */ __RPC__out METRO_MONITOR_MODE *pMode) = 0;

	virtual HRESULT STDMETHODCALLTYPE IsLauncherVisible(
		/* [out] */ __RPC__out BOOL *pfVisible) = 0;

	virtual HRESULT STDMETHODCALLTYPE Advise(
		/* [in] */ __RPC__in_opt IMonitorModeEvents *pCallback,
		/* [out] */ __RPC__out DWORD *pdwCookie) = 0;

	virtual HRESULT STDMETHODCALLTYPE Unadvise(
		/* [in] */ DWORD dwCookie) = 0;
};

static const CLSID CLSID_MetroMode={0x7E5FE3D9,0x985F,0x4908,{0x91, 0xF9, 0xEE, 0x19, 0xF9, 0xFD, 0x15, 0x14}};

#endif

static bool IsMetroMode( void )
{
	CComPtr<IMetroMode> pMetroMode;
	pMetroMode.CoCreateInstance(CLSID_MetroMode);
	BOOL bMetro;
	return (pMetroMode && SUCCEEDED(pMetroMode->IsLauncherVisible(&bMetro)) && bMetro);
}

static LRESULT CALLBACK HookAppManager( int code, WPARAM wParam, LPARAM lParam )
{
	if (code==HC_ACTION)
	{
		MSG *msg=(MSG*)lParam;
		if ((msg->message==WM_MOUSEMOVE || msg->message==WM_LBUTTONDOWN) && GetSettingBool(L"DisableHotCorner") && !IsMetroMode())
		{
			// ignore the mouse messages if there is a menu
			GUITHREADINFO info={sizeof(info)};
			if (GetGUIThreadInfo(GetCurrentThreadId(),&info) && (info.flags&GUI_INMENUMODE))
				return CallNextHookEx(NULL,code,wParam,lParam);
			APPBARDATA appbar={sizeof(appbar),g_TaskBar};
			SHAppBarMessage(ABM_GETTASKBARPOS,&appbar);
			if (appbar.uEdge==ABE_BOTTOM)
			{
				// check if the mouse is over the taskbar
				RECT rc;
				GetWindowRect(g_TaskBar,&rc);
				CPoint pt(GetMessagePos());
				if (PtInRect(&rc,pt))
				{
					if (msg->message==WM_LBUTTONDOWN)
					{
						// forward the mouse click to the taskbar
						PostMessage(g_TaskBar,WM_NCLBUTTONDOWN,MK_LBUTTON,MAKELONG(pt.x,pt.y));
						msg->message=WM_NULL;
					}
					wchar_t className[256]={0};
					GetClassName(msg->hwnd,className,_countof(className));
					if (wcscmp(className,L"ImmersiveSwitchList")==0)
					{
						// suppress the opening of the ImmersiveSwitchList
						msg->message=WM_NULL;
						ShowWindow(msg->hwnd,SW_HIDE); // hide the popup
						g_SwitchList=msg->hwnd;
					}
				}
			}
		}
	}
	return CallNextHookEx(NULL,code,wParam,lParam);
}

static LRESULT CALLBACK HookKeyboardLL( int code, WPARAM wParam, LPARAM lParam )
{
	static bool s_bLWinPressed, s_bRWinPressed;
	if (wParam==WM_KEYDOWN)
	{
		KBDLLHOOKSTRUCT *pKbd=(KBDLLHOOKSTRUCT*)lParam;
		s_bLWinPressed=(pKbd->vkCode==VK_LWIN);
		s_bRWinPressed=(pKbd->vkCode==VK_RWIN);
	}
	if (wParam==WM_KEYUP)
	{
		KBDLLHOOKSTRUCT *pKbd=(KBDLLHOOKSTRUCT*)lParam;
		if (((pKbd->vkCode==VK_LWIN && s_bLWinPressed) || (pKbd->vkCode==VK_RWIN && s_bRWinPressed)) && GetAsyncKeyState(VK_SHIFT)<0)
		{
			PostMessage(g_TaskBar,g_StartMenuMsg,MSG_SHIFTWIN,0);
		}
	}
	return CallNextHookEx(NULL,code,wParam,lParam);
}

// Set the hotkeys and controls for the start menu
void EnableHotkeys( THotkeys enable )
{
	if (!g_TaskBar)
		return;
	if (GetWindowThreadProcessId(g_TaskBar,NULL)!=GetCurrentThreadId())
	{
		PostMessage(g_TaskBar,g_StartMenuMsg,MSG_HOTKEYS,enable);
		return;
	}

	// must be executed in the same thread as the start button (otherwise RegisterHotKey doesn't work). also prevents race conditions
	bool bHook=(enable==HOTKEYS_SETTINGS || (enable==HOTKEYS_NORMAL && GetSettingInt(L"ShiftWin")!=0));
	if (bHook)
	{
		if (!g_KeyboardHook)
			g_KeyboardHook=SetWindowsHookEx(WH_KEYBOARD_LL,HookKeyboardLL,g_Instance,NULL);
	}
	else
	{
		if (g_KeyboardHook)
			UnhookWindowsHookEx(g_KeyboardHook);
		g_KeyboardHook=NULL;
	}

	if (g_HotkeyCSM)
		UnregisterHotKey(g_TaskBar,g_HotkeyCSMID);
	g_HotkeyCSM=0;

	if (g_HotkeyWSM)
		UnregisterHotKey(g_TaskBar,g_HotkeyWSMID);
	g_HotkeyWSM=0;

	if (enable==HOTKEYS_NORMAL)
	{
		g_HotkeyCSM=GetSettingInt(L"CSMHotkey");
		if (g_HotkeyCSM)
		{
			int mod=0;
			if (g_HotkeyCSM&(HOTKEYF_SHIFT<<8)) mod|=MOD_SHIFT;
			if (g_HotkeyCSM&(HOTKEYF_CONTROL<<8)) mod|=MOD_CONTROL;
			if (g_HotkeyCSM&(HOTKEYF_ALT<<8)) mod|=MOD_ALT;
			BOOL err=RegisterHotKey(g_TaskBar,g_HotkeyCSMID,mod,g_HotkeyCSM&255);
			int q=0;
		}

		g_HotkeyWSM=GetSettingInt(L"WSMHotkey");
		if (g_HotkeyWSM)
		{
			int mod=0;
			if (g_HotkeyWSM&(HOTKEYF_SHIFT<<8)) mod|=MOD_SHIFT;
			if (g_HotkeyWSM&(HOTKEYF_CONTROL<<8)) mod|=MOD_CONTROL;
			if (g_HotkeyWSM&(HOTKEYF_ALT<<8)) mod|=MOD_ALT;
			RegisterHotKey(g_TaskBar,g_HotkeyWSMID,mod,g_HotkeyWSM&255);
		}
	}
}

static LRESULT CALLBACK SubclassUserPicProc( HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData )
{
	if (uMsg==WM_WINDOWPOSCHANGING && !(((WINDOWPOS*)lParam)->flags&SWP_NOMOVE))
	{
		if (GetSettingBool(L"HideUserPic"))
		{
			((WINDOWPOS*)lParam)->x=-32000;
			((WINDOWPOS*)lParam)->y=-32000;
		}
	}
	return DefSubclassProc(hWnd,uMsg,wParam,lParam);
}

static LRESULT CALLBACK SubclassTopMenuProc( HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData )
{
	if (uMsg==WM_ACTIVATE && GetSettingBool(L"CascadeAll"))
	{
		if (!wParam)
		{
			if (CMenuContainer::s_pDragSource) return 0;
			// check if another menu window is being activated
			// if not, close all menus
			for (std::vector<CMenuContainer*>::const_iterator it=CMenuContainer::s_Menus.begin();it!=CMenuContainer::s_Menus.end();++it)
				if ((*it)->m_hWnd==(HWND)lParam)
					return 0;
		}
	}
	if (uMsg==WM_WINDOWPOSCHANGED && (((WINDOWPOS*)lParam)->flags&SWP_SHOWWINDOW))
	{
		g_LastHoverPos=GetMessagePos();
		if (GetSettingInt(L"InitiallySelect")==1)
			PostMessage(hWnd,WM_CLEAR,'CLSH',0);
		PressStartButton(true);
	}
	if (uMsg==WM_CLEAR && wParam=='CLSH')
	{
		SetFocus(g_ProgramsButton);
		return 0;
	}
	if (uMsg==WM_SHOWWINDOW)
	{
		if (!wParam)
		{
			CMenuContainer::CloseProgramsMenu();
			PressStartButton(false);
		}
		g_bAllProgramsTimer=false;
		KillTimer(g_ProgramsButton,'CLSM');
	}
	if (uMsg==WM_DESTROY)
		g_TopMenu=NULL;
	if (uMsg==WM_ACTIVATEAPP && !wParam)
	{
		if (CMenuContainer::s_pDragSource) return 0;
	}
	if (uMsg==WM_MOUSEACTIVATE && GetSettingBool(L"CascadeAll"))
	{
		CPoint pt(GetMessagePos());
		if (WindowFromPoint(pt)==g_ProgramsButton && CMenuContainer::IsMenuOpened())
			return MA_NOACTIVATEANDEAT;
		CMenuContainer::CloseProgramsMenu();
	}
	return DefSubclassProc(hWnd,uMsg,wParam,lParam);
}

static LRESULT CALLBACK SubclassProgramsProc( HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData )
{
	if (uMsg==WM_COMMAND && wParam==IDOK && GetSettingBool(L"CascadeAll"))
	{
		if (GetKeyState(VK_SHIFT)<0)
		{
			if (CMenuContainer::IsMenuOpened())
				return 0; // ignore shift+click when the menu is opened
		}
		else
		{
			if (!CMenuContainer::IsMenuOpened())
				CMenuContainer::ToggleStartMenu(hWnd,GetKeyState(VK_SPACE)<0 || GetKeyState(VK_RETURN)<0 || GetKeyState(VK_LEFT)<0 || GetKeyState(VK_RIGHT)<0,true);
			return 0;
		}
	}
	if (uMsg==WM_DRAWITEM && wParam==IDOK && CMenuContainer::IsMenuOpened())
	{
		DRAWITEMSTRUCT *pDraw=(DRAWITEMSTRUCT*)lParam;
		pDraw->itemState=ODS_HOTLIGHT; // draw highlighted when the menu is open
	}
	return DefSubclassProc(hWnd,uMsg,wParam,lParam);
}

static BOOL CALLBACK FindWindowsMenuProc( HWND hwnd, LPARAM lParam )
{
	wchar_t name[100];
	GetClassName(hwnd,name,_countof(name));
	if (_wcsicmp(name,L"DV2ControlHost")==0)
	{
		HWND w1=hwnd;
		if (LOWORD(GetVersion())==0x0006)
		{
			w1=FindWindowEx(w1,NULL,L"Desktop Open Pane Host",NULL);
			if (!w1) return TRUE;
		}
		w1=FindWindowEx(w1,NULL,L"Desktop More Programs Pane",NULL);
		if (!w1) return TRUE;
		HWND w2=GetDlgItem(w1,IDOK);
		if (!w2) return TRUE;

		g_TopMenu=hwnd;
		g_AllPrograms=w1;
		g_ProgramsButton=w2;
		return FALSE;
	}
	return TRUE;
}

static void FindWindowsMenu( void )
{
	if (!g_TopMenu)
	{
		ATLASSERT(GetCurrentThreadId()==GetWindowThreadProcessId(g_TaskBar,NULL));
		EnumThreadWindows(GetCurrentThreadId(),FindWindowsMenuProc,0);
		if (g_TopMenu)
		{
			g_UserPic=FindWindow(L"Desktop User Picture",NULL);
			SetWindowSubclass(g_UserPic,SubclassUserPicProc,'CLSH',0);
			SetWindowSubclass(g_TopMenu,SubclassTopMenuProc,'CLSH',0);
			SetWindowSubclass(g_AllPrograms,SubclassProgramsProc,'CLSH',0);
		}
	}
}

static LRESULT CALLBACK SubclassTaskBarProc( HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData )
{
	if (uMsg==WM_MOUSEACTIVATE && HIWORD(lParam)==WM_MBUTTONDOWN)
	{
		FindWindowsMenu();
		if (g_TopMenu && IsWindowVisible(g_TopMenu))
		{
			DefSubclassProc(hWnd,uMsg,wParam,lParam);
			return MA_ACTIVATEANDEAT; // ignore the next middle click, so it doesn't re-open the start menu
		}
	}
	if (g_bReplaceButton)
	{
		if ((uMsg==WM_NCMOUSEMOVE || uMsg==WM_MOUSEMOVE) && PointAroundStartButton())
		{
			TaskBarMouseMove();
		}
		if (uMsg==WM_WINDOWPOSCHANGED)
		{
			if (IsStartButtonSmallIcons()!=IsTaskbarSmallIcons())
				RecreateStartButton();

			WINDOWPOS *pPos=(WINDOWPOS*)lParam;
			RECT rcTask;
			GetWindowRect(hWnd,&rcTask);
			APPBARDATA appbar={sizeof(appbar),hWnd};
			SHAppBarMessage(ABM_GETTASKBARPOS,&appbar);
			DWORD flags=SWP_NOACTIVATE|SWP_NOOWNERZORDER|SWP_SHOWWINDOW|SWP_NOSIZE;
			APPBARDATA appbar2={sizeof(appbar2)};
			
			MONITORINFO info={sizeof(info)};
			GetMonitorInfo(MonitorFromWindow(g_TaskBar,MONITOR_DEFAULTTONEAREST),&info);
			if (SHAppBarMessage(ABM_GETSTATE,&appbar2)&ABS_AUTOHIDE)
			{
				bool bHide=false;
				if (appbar.uEdge==ABE_LEFT)
					bHide=(rcTask.right<info.rcMonitor.left+5);
				else if (appbar.uEdge==ABE_RIGHT)
					bHide=(rcTask.left>info.rcMonitor.right-5);
				else if (appbar.uEdge==ABE_TOP)
					bHide=(rcTask.bottom<info.rcMonitor.top+5);
				else
					bHide=(rcTask.top>info.rcMonitor.bottom-5);
				if (bHide)
					flags=(flags&~SWP_SHOWWINDOW)|SWP_HIDEWINDOW;
			}
			DWORD version=LOWORD(GetVersion());
			version=MAKEWORD(HIBYTE(version),LOBYTE(version));
			bool bClassic;
			if (version<0x0602)
				bClassic=!IsAppThemed();
			else
			{
				HIGHCONTRAST contrast={sizeof(contrast)};
				bClassic=(SystemParametersInfo(SPI_GETHIGHCONTRAST,sizeof(contrast),&contrast,0) && (contrast.dwFlags&HCF_HIGHCONTRASTON));
			}
			if (!bClassic)
			{
				if (appbar.uEdge==ABE_TOP)
					OffsetRect(&rcTask,0,-1);
				else if (appbar.uEdge==ABE_BOTTOM)
					OffsetRect(&rcTask,0,1);
			}
			HWND zPos=(pPos->flags&SWP_NOZORDER)?HWND_TOPMOST:pPos->hwndInsertAfter;
			if (zPos==HWND_TOP && !(GetWindowLong(g_StartButton,GWL_EXSTYLE)&WS_EX_TOPMOST))
				zPos=HWND_TOPMOST;

			SIZE size=GetStartButtonSize();
			int x, y;
			if (GetWindowLong(g_Rebar,GWL_STYLE)&CCS_VERT)
			{
				x=(rcTask.left+rcTask.right-size.cx)/2;
				y=rcTask.top;
			}
			else if (GetWindowLong(g_Rebar,GWL_EXSTYLE)&WS_EX_LAYOUTRTL)
			{
				x=rcTask.right-size.cx;
				y=(rcTask.top+rcTask.bottom-size.cy)/2;
			}
			else
			{
				x=rcTask.left;
				y=(rcTask.top+rcTask.bottom-size.cy)/2;
			}
			RECT rcButton={x,y,x+size.cx,y+size.cy};
			RECT rc;
			IntersectRect(&rc,&rcButton,&info.rcMonitor);
			HRGN rgn=CreateRectRgn(rc.left-x,rc.top-y,rc.right-x,rc.bottom-y);
			if (!SetWindowRgn(g_StartButton,rgn,FALSE))
				DeleteObject(rgn);
			SetWindowPos(g_StartButton,zPos,x,y,0,0,flags);
		}
		if (uMsg==WM_THEMECHANGED)
		{
			RecreateStartButton();
		}
		if (uMsg==WM_PAINT && g_StartButtonOld && !IsAppThemed())
		{
			PAINTSTRUCT ps;
			HDC hdc=BeginPaint(hWnd,&ps);
			EndPaint(hWnd,&ps);
			return 0;
		}
	}
	if (uMsg==WM_TIMER && wParam=='CLSM')
	{
		if (!IsMetroMode())
		{
			KillTimer(hWnd,'CLSM');
			return 0;
		}
		SetForegroundWindow(hWnd);
		INPUT inputs[4]={
			{INPUT_KEYBOARD},
			{INPUT_KEYBOARD},
			{INPUT_KEYBOARD},
			{INPUT_KEYBOARD},
		};
		inputs[0].ki.wVk=VK_LWIN;
		inputs[1].ki.wVk='D';
		inputs[2].ki.wVk='D';
		inputs[2].ki.dwFlags=KEYEVENTF_KEYUP;
		inputs[3].ki.wVk=VK_LWIN;
		inputs[3].ki.dwFlags=KEYEVENTF_KEYUP;
		SendInput(_countof(inputs),inputs,sizeof(INPUT));
		g_MetroModeCount--;
		if (g_MetroModeCount<=0)
			KillTimer(hWnd,'CLSM');
		return 0;
	}
	return DefSubclassProc(hWnd,uMsg,wParam,lParam);
}

static LRESULT CALLBACK SubclassRebarProc( HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData )
{
	if (uMsg==WM_WINDOWPOSCHANGING)
	{
		WINDOWPOS *pPos=(WINDOWPOS*)lParam;
		if (!(pPos->flags&SWP_NOMOVE) || !(pPos->flags&SWP_NOSIZE))
		{
			if (pPos->flags&(SWP_NOMOVE|SWP_NOSIZE))
			{
				RECT rc;
				GetWindowRect(hWnd,&rc);
				MapWindowPoints(NULL,GetParent(hWnd),(POINT*)&rc,2);
				if (pPos->flags&SWP_NOMOVE)
				{
					pPos->x=rc.left;
					pPos->y=rc.top;
				}
				else
				{
					pPos->cx=rc.right-rc.left;
					pPos->cy=rc.bottom-rc.top;
				}
			}
			int dx=0, dy=0;
			if (GetWindowLong(hWnd,GWL_STYLE)&CCS_VERT)
			{
				dy=g_RebarOffset.cy-pPos->y;
			}
			else
			{
				dx=g_RebarOffset.cx-pPos->x;
			}
			if (dx || dy)
			{
				pPos->x+=dx;
				pPos->cx-=dx;
				pPos->y+=dy;
				pPos->cy-=dy;
				pPos->flags&=~(SWP_NOMOVE|SWP_NOSIZE);
			}
		}
	}
	return DefSubclassProc(hWnd,uMsg,wParam,lParam);
}

static void InitStartMenuDLL( void )
{
	int level=GetSettingInt(L"CrashDump");
	if (level>=1 && level<=3)
	{
		if (level==1) MiniDumpType=MiniDumpNormal;
		if (level==2) MiniDumpType=MiniDumpWithDataSegs;
		if (level==3) MiniDumpType=MiniDumpWithFullMemory;
		SetUnhandledExceptionFilter(TopLevelFilter);
		g_bCrashDump=true;
	}
	FindTaskBar();
	g_ProgWin=FindWindowEx(NULL,NULL,L"Progman",NULL);
	DWORD thread=GetWindowThreadProcessId(g_ProgWin,NULL);
	g_ProgHook=SetWindowsHookEx(WH_GETMESSAGE,HookProgMan,NULL,thread);
	g_StartHook=SetWindowsHookEx(WH_GETMESSAGE,HookStartButton,NULL,GetCurrentThreadId());
	HWND hwnd=FindWindow(L"ClassicStartMenu.CStartHookWindow",L"StartHookWindow");
	LoadLibrary(L"ClassicStartMenuDLL.dll"); // keep the DLL from unloading
	if (hwnd) PostMessage(hwnd,WM_CLEAR,0,0); // tell the exe to unhook this hook
	g_bReplaceButton=GetSettingBool(L"EnableStartButton");
	if (g_bReplaceButton)
	{
		g_Rebar=FindWindowEx(g_TaskBar,NULL,REBARCLASSNAME,NULL);
		if (g_Rebar)
		{
			g_StartButtonOld=g_StartButton;
			if (g_StartButtonOld)
			{
				ShowWindow(g_StartButtonOld,SW_HIDE);
				if (LOWORD(GetVersion())==0x0106)
				{
					// Windows 7 draws the start button on the taskbar as well
					// so we zero out the bitmap resources
					HMODULE hExplorer=GetModuleHandle(NULL);
					for (int res=0;res<_countof(g_StartButtonOldSizes);res++)
					{
						HRSRC hrSrc=FindResource(hExplorer,MAKEINTRESOURCE(res+FIRST_BUTTON_BITMAP),RT_BITMAP);
						if (hrSrc)
						{
							HGLOBAL hRes=LoadResource(hExplorer,hrSrc);
							if (hRes)
							{
								void *pRes=LockResource(hRes);
								if (pRes)
								{
									DWORD old;
									BITMAPINFOHEADER *pHeader=(BITMAPINFOHEADER*)pRes;
									if (pHeader->biWidth)
									{
										g_StartButtonOldSizes[res]=MAKELONG(pHeader->biWidth,pHeader->biHeight);
										VirtualProtect(pRes,sizeof(BITMAPINFOHEADER),PAGE_READWRITE,&old);
										pHeader->biHeight=pHeader->biWidth=0;
										VirtualProtect(pRes,sizeof(BITMAPINFOHEADER),old,&old);
									}
								}
							}
						}
					}
				}
				SendMessage(g_TaskBar,WM_SETTINGCHANGE,0,0);
			}
			SetWindowSubclass(g_TaskBar,SubclassTaskBarProc,'CLSH',0);
			SetWindowSubclass(g_Rebar,SubclassRebarProc,'CLSH',0);
			g_StartButton=NULL;
			g_bStartScreen=false;
			RecreateStartButton();
		}
		else
			g_bReplaceButton=false;
	}
	else
	{
		SetWindowSubclass(g_TaskBar,SubclassTaskBarProc,'CLSH',0);
	}
#ifdef HOOK_DROPTARGET
	if (g_StartButton)
	{
		if (!g_bReplaceButton)
		{
			g_pOriginalTarget=(IDropTarget*)GetProp(g_StartButton,L"OleDropTargetInterface");
			if (g_pOriginalTarget)
				RevokeDragDrop(g_StartButton);
		}
		CStartMenuTarget *pNewTarget=new CStartMenuTarget();
		RegisterDragDrop(g_StartButton,pNewTarget);
		pNewTarget->Release();
	}
#endif
	if (g_bReplaceButton)
	{
		HWND appManager=FindWindow(L"ApplicationManager_DesktopShellWindow",NULL);
		if (appManager)
		{
			DWORD thread=GetWindowThreadProcessId(appManager,NULL);
			g_AppManagerHook=SetWindowsHookEx(WH_GETMESSAGE,HookAppManager,g_Instance,thread);
		}
		if (GetSettingBool(L"SkipMetro"))
		{
			g_MetroModeCount=GetSettingInt(L"SkipMetroCount");
			SetTimer(g_TaskBar,'CLSM',500,NULL);
		}
	}
}

void RecreateStartButton( void )
{
	if (!g_bReplaceButton) return;
	RevokeDragDrop(g_StartButton);
	DestroyStartButton();
	RECT rcTask;
	GetWindowRect(g_TaskBar,&rcTask);
	RECT rcTask2=rcTask;
	APPBARDATA appbar={sizeof(appbar),g_TaskBar};
	SHAppBarMessage(ABM_GETTASKBARPOS,&appbar);
	DWORD version=LOWORD(GetVersion());
	version=MAKEWORD(HIBYTE(version),LOBYTE(version));
	bool bClassic;
	if (version<0x0602)
		bClassic=!IsAppThemed();
	else
	{
		HIGHCONTRAST contrast={sizeof(contrast)};
		bClassic=(SystemParametersInfo(SPI_GETHIGHCONTRAST,sizeof(contrast),&contrast,0) && (contrast.dwFlags&HCF_HIGHCONTRASTON));
	}
	if (!bClassic)
	{
		if (appbar.uEdge==ABE_TOP)
			OffsetRect(&rcTask2,0,-1);
		else if (appbar.uEdge==ABE_BOTTOM)
			OffsetRect(&rcTask2,0,1);
	}
	g_StartButton=CreateStartButton(g_TaskBar,g_Rebar,rcTask2);
	CStartMenuTarget *pNewTarget=new CStartMenuTarget();
	RegisterDragDrop(g_StartButton,pNewTarget);
	pNewTarget->Release();

	g_RebarOffset=GetStartButtonSize();

	PostMessage(g_TaskBar,WM_SIZE,SIZE_RESTORED,MAKELONG(rcTask.right-rcTask.left,rcTask.bottom-rcTask.top));
}

static DWORD WINAPI ExitThreadProc( void *param )
{
	Sleep(1000); // wait a second! hopefully by then the hooks will be finished and no more of our code will be executing
	// send WM_CLOSE to the window in ClassicStartMenu.exe to close that process
	if (param) PostMessage((HWND)param,WM_CLOSE,0,0);
	FreeLibraryAndExitThread(g_Instance,0);
}

static void CleanStartMenuDLL( void )
{
	// cleanup
	if (g_Owner.m_hWnd) g_Owner.DestroyWindow();
	CloseSettings();
	CMenuContainer::CloseStartMenu();
	CMenuFader::ClearAll();
	g_IconManager.StopLoading();
	UnhookDropTarget();
	EnableHotkeys(HOTKEYS_CLEAR);
	HWND hwnd=FindWindow(L"ClassicStartMenu.CStartHookWindow",L"StartHookWindow");
	UnhookWindowsHookEx(g_ProgHook);
	UnhookWindowsHookEx(g_StartHook);
	if (g_AppManagerHook) UnhookWindowsHookEx(g_AppManagerHook);
	g_AppManagerHook=NULL;
	RemoveWindowSubclass(g_TaskBar,SubclassTaskBarProc,'CLSH');
	if (g_StartButtonOld)
	{
		if (LOWORD(GetVersion())==0x0106)
		{
			// restore the bitmap sizes
			HMODULE hExplorer=GetModuleHandle(NULL);
			for (int res=0;res<_countof(g_StartButtonOldSizes);res++)
			{
				HRSRC hrSrc=FindResource(hExplorer,MAKEINTRESOURCE(res+FIRST_BUTTON_BITMAP),RT_BITMAP);
				if (hrSrc)
				{
					HGLOBAL hRes=LoadResource(hExplorer,hrSrc);
					if (hRes)
					{
						void *pRes=LockResource(hRes);
						if (pRes)
						{
							DWORD old;
							BITMAPINFOHEADER *pHeader=(BITMAPINFOHEADER*)pRes;
							if (g_StartButtonOldSizes[res])
							{
								VirtualProtect(pRes,sizeof(BITMAPINFOHEADER),PAGE_READWRITE,&old);
								pHeader->biWidth=LOWORD(g_StartButtonOldSizes[res]);
								pHeader->biHeight=HIWORD(g_StartButtonOldSizes[res]);
								VirtualProtect(pRes,sizeof(BITMAPINFOHEADER),old,&old);
							}
						}
					}
				}
			}
		}
		SendMessage(g_TaskBar,WM_SETTINGCHANGE,0,0);
		SendMessage(g_StartButtonOld,WM_THEMECHANGED,0,0);
		ShowWindow(g_StartButtonOld,SW_SHOW);
	}
	g_StartButtonOld=NULL;
	if (g_Rebar)
	{
		RemoveWindowSubclass(g_Rebar,SubclassRebarProc,'CLSH');
		RECT rc;
		GetWindowRect(g_TaskBar,&rc);
		PostMessage(g_TaskBar,WM_SIZE,SIZE_RESTORED,MAKELONG(rc.right-rc.left,rc.bottom-rc.top));
		if (g_bReplaceButton)
			DestroyStartButton();
	}
	if (g_TopMenu)
	{
		RemoveWindowSubclass(g_UserPic,SubclassUserPicProc,'CLSH');
		RemoveWindowSubclass(g_TopMenu,SubclassTopMenuProc,'CLSH');
		RemoveWindowSubclass(g_AllPrograms,SubclassProgramsProc,'CLSH');
	}
	if (g_bCrashDump)
	{
		SetUnhandledExceptionFilter(NULL);
		g_bCrashDump=false;
	}
	if (g_bStartButtonTimer)
	{
		g_bStartButtonTimer=false;
		KillTimer(g_StartButton,'CLSM');
	}
	g_StartButton=NULL;
//	g_TaskBar=NULL; - do not clear g_TaskBar to prevent InitStartMenuDLL from being called again
	g_Rebar=NULL;
	g_TopMenu=NULL;
	g_bReplaceButton=false;

	// we need to unload the DLL here. but we can't just call FreeLibrary because it will unload the code
	// while it is still executing. So we create a separate thread and use FreeLibraryAndExitThread
	CreateThread(NULL,0,ExitThreadProc,(void*)hwnd,0,NULL);
}

///////////////////////////////////////////////////////////////////////////////

// WH_GETMESSAGE hook for the Progman window
STARTMENUAPI LRESULT CALLBACK HookProgMan( int code, WPARAM wParam, LPARAM lParam )
{
	if (code==HC_ACTION)
	{
		MSG *msg=(MSG*)lParam;
		if (msg->message==WM_SYSCOMMAND && (msg->wParam&0xFFF0)==SC_TASKLIST && msg->lParam!='CLSM')
		{
			// Win button pressed
			FindTaskBar();
			int control=GetSettingInt(L"WinKey");
			if (control!=OPEN_WINDOWS)
			{
				msg->message=WM_NULL;
				if (control==OPEN_CLASSIC)
					PostMessage(g_TaskBar,g_StartMenuMsg,MSG_TOGGLE,0);
			}
			else
				PostMessage(g_TaskBar,g_StartMenuMsg,MSG_FINDMENU,0);
		}
	}
	return CallNextHookEx(NULL,code,wParam,lParam);
}

static bool WindowsMenuOpened( void )
{
	if (g_bReplaceButton && !g_StartButtonOld)
		return IsMetroMode();
	else
		return (SendMessage(g_StartButton,BM_GETSTATE,0,0)&BST_PUSHED)!=0;
}

// WH_GETMESSAGE hook for the start button window
STARTMENUAPI LRESULT CALLBACK HookStartButton( int code, WPARAM wParam, LPARAM lParam )
{
	if (code==HC_ACTION && !g_bInMenu)
	{
		MSG *msg=(MSG*)lParam;
		FindTaskBar();
		if (IsSettingsMessage(msg))
		{
			msg->message=WM_NULL;
			return 0;
		}
		if (((msg->message>=WM_MOUSEFIRST && msg->message<=WM_MOUSELAST) || msg->message==WM_MOUSEHOVER || msg->message==WM_MOUSELEAVE) && CMenuContainer::ProcessMouseMessage(msg->hwnd,msg->message,msg->wParam,msg->lParam))
		{
			msg->message=WM_NULL;
			return 0;
		}
		if (msg->message==g_StartMenuMsg && msg->hwnd==g_TaskBar)
		{
			msg->message=WM_NULL;
			FindWindowsMenu();
			if (msg->wParam==MSG_TOGGLE || (msg->wParam==MSG_OPEN && !CMenuContainer::IsMenuOpened()))
				ToggleStartMenu(g_StartButton,true);
			else if (msg->wParam==MSG_SETTINGS)
			{
				if (GetSettingBool(L"EnableSettings"))
					EditSettings(false);
			}
			else if (msg->wParam==MSG_SHIFTWIN)
			{
				int control=GetSettingInt(L"ShiftWin");
				if (control==OPEN_CLASSIC)
					ToggleStartMenu(g_StartButton,true);
				else if (control==OPEN_WINDOWS)
					PostMessage(g_ProgWin,WM_SYSCOMMAND,SC_TASKLIST,'CLSM');
			}
			else if (msg->wParam==MSG_DRAG || msg->wParam==MSG_SHIFTDRAG)
			{
				int control=GetSettingInt((msg->wParam==MSG_DRAG)?L"MouseClick":L"ShiftClick");
				if (control==OPEN_CLASSIC)
					ToggleStartMenu(g_StartButton,true);
				else if (control==OPEN_WINDOWS)
					PostMessage(g_ProgWin,WM_SYSCOMMAND,SC_TASKLIST,'CLSM');
			}
			else if (msg->wParam==MSG_EXIT)
			{
				LRESULT res=CallNextHookEx(NULL,code,wParam,lParam);
				CleanStartMenuDLL();
				return res; // we should exit as quickly as possible now. the DLL is about to be unloaded
			}
			else if (msg->wParam==MSG_HOTKEYS)
			{
				EnableHotkeys((THotkeys)msg->lParam);
			}
		}

		if (msg->message==WM_HOTKEY && msg->hwnd==g_TaskBar)
		{
			if (msg->wParam==g_HotkeyCSMID)
			{
				msg->message=WM_NULL;
				if (g_StartButton)
					SetForegroundWindow(g_StartButton);
				ToggleStartMenu(g_StartButton,true);
			}
			else if (msg->wParam==g_HotkeyWSMID)
				PostMessage(g_ProgWin,WM_SYSCOMMAND,SC_TASKLIST,'CLSM');
		}

		if (msg->message==WM_KEYDOWN && msg->hwnd==g_TaskBar && (msg->wParam==VK_SPACE || msg->wParam==VK_RETURN))
		{
			GUITHREADINFO info={sizeof(info)};
			if (!GetGUIThreadInfo(GetCurrentThreadId(),&info) || !(info.flags&GUI_INMENUMODE))
			{
				FindWindowsMenu();
				if (GetSettingInt(L"WinKey")==OPEN_CLASSIC)
				{
					msg->message=WM_NULL;
					if (g_StartButton)
						SetForegroundWindow(g_StartButton);
					ToggleStartMenu(g_StartButton,true);
				}
			}
		}

		if (msg->message==WM_KEYDOWN && msg->wParam==VK_TAB && CMenuContainer::IsMenuWindow(msg->hwnd))
		{
			// the taskbar steals the Tab key. we need to forward it to the menu instead
			SendMessage(msg->hwnd,msg->message,msg->wParam,msg->lParam);
			msg->message=WM_NULL;
		}

		if (msg->message==WM_SYSKEYDOWN && msg->wParam==VK_RETURN && CMenuContainer::IsMenuWindow(msg->hwnd))
		{
			// the taskbar steals the Alt+Enter key. we need to forward it to the menu instead
			SendMessage(msg->hwnd,msg->message,msg->wParam,msg->lParam);
			msg->message=WM_NULL;
		}

		if ((msg->message==WM_LBUTTONDOWN || msg->message==WM_LBUTTONDBLCLK || msg->message==WM_MBUTTONDOWN || msg->message==WM_MBUTTONDBLCLK) && msg->hwnd==g_StartButton)
		{
			FindWindowsMenu();
			const wchar_t *name;
			if (msg->message==WM_MBUTTONDOWN || msg->message==WM_MBUTTONDBLCLK)
				name=L"MiddleClick";
			else if (GetKeyState(VK_CONTROL)<0)
				name=L"Hover";
			else if (GetKeyState(VK_SHIFT)<0)
				name=L"ShiftClick";
			else
				name=L"MouseClick";
			int control=GetSettingInt(name);
			if (control==OPEN_NOTHING)
				msg->message=WM_NULL;
			if (control==OPEN_CLASSIC)
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
			else if (control==OPEN_WINDOWS && (msg->message==WM_MBUTTONDOWN || msg->message==WM_MBUTTONDBLCLK || g_bReplaceButton))
				PostMessage(g_ProgWin,WM_SYSCOMMAND,SC_TASKLIST,'CLSM');
		}

		if (g_bReplaceButton)
		{
			if (msg->hwnd==g_TaskBar && (msg->message==WM_NCLBUTTONDOWN || msg->message==WM_NCLBUTTONDBLCLK || msg->message==WM_NCMBUTTONDOWN || msg->message==WM_NCMBUTTONDBLCLK || msg->message==WM_LBUTTONDOWN || msg->message==WM_LBUTTONDBLCLK || msg->message==WM_MBUTTONDOWN || msg->message==WM_MBUTTONDBLCLK))
			{
				if (PointAroundStartButton())
				{
					const wchar_t *name;
					if (msg->message==WM_NCMBUTTONDOWN || msg->message==WM_MBUTTONDOWN || msg->message==WM_NCMBUTTONDBLCLK || msg->message==WM_MBUTTONDBLCLK)
						name=L"MiddleClick";
					else if (GetKeyState(VK_SHIFT)<0)
						name=L"ShiftClick";
					else
						name=L"MouseClick";
					int control=GetSettingInt(name);
					if (control==OPEN_CLASSIC)
					{
						DWORD keyboard;
						SystemParametersInfo(SPI_GETKEYBOARDCUES,NULL,&keyboard,0);
						ToggleStartMenu(g_StartButton,keyboard!=0);
					}
					else if (control==OPEN_WINDOWS)
						PostMessage(g_ProgWin,WM_SYSCOMMAND,SC_TASKLIST,'CLSM');
					msg->message=WM_NULL;
				}
			}
		}
		else
		{
			if ((msg->message==WM_NCLBUTTONDOWN || msg->message==WM_NCLBUTTONDBLCLK || msg->message==WM_NCMBUTTONDOWN || msg->message==WM_NCMBUTTONDBLCLK) && msg->hwnd==g_TaskBar
				&& (msg->wParam==HTCAPTION || !IsAppThemed())) // HACK: in Classic mode the start menu can show up even if wParam is not HTCAPTION (most likely a bug in Windows)
			{
				FindWindowsMenu();
				const wchar_t *name;
				if (msg->message==WM_NCMBUTTONDOWN || msg->message==WM_NCMBUTTONDBLCLK)
					name=L"MiddleClick";
				else if (GetKeyState(VK_SHIFT)<0)
					name=L"ShiftClick";
				else
					name=L"MouseClick";
				int control=GetSettingInt(name);
				if (control==OPEN_NOTHING)
					msg->message=WM_NULL;
				else if (control==OPEN_CLASSIC)
				{
					if (PointAroundStartButton())
					{
						// click on the taskbar around the start menu - toggle the menu
						DWORD keyboard;
						SystemParametersInfo(SPI_GETKEYBOARDCUES,NULL,&keyboard,0);
						ToggleStartMenu(g_StartButton,keyboard!=0);
						msg->message=WM_NULL;
					}
				}
				else if (control==OPEN_WINDOWS && (msg->message==WM_NCMBUTTONDOWN || msg->message==WM_NCMBUTTONDBLCLK))
					PostMessage(g_ProgWin,WM_SYSCOMMAND,SC_TASKLIST,'CLSM');
			}
		}

		if (msg->message==WM_TIMER && msg->hwnd==g_TaskBar && CMenuContainer::IgnoreTaskbarTimers())
		{
			// stop the taskbar timer messages. prevents the auto-hide taskbar from closing
			msg->message=WM_NULL;
		}

		if (msg->message==WM_MOUSEMOVE && msg->hwnd==g_ProgramsButton && GetSettingBool(L"CascadeAll") && !(msg->wParam&MK_SHIFT))
		{
			DWORD pos=GetMessagePos();
			if (pos!=g_LastHoverPos && !g_bAllProgramsTimer)
			{
				g_bAllProgramsTimer=true;
				bool bDef;
				DWORD time=GetSettingInt(L"AllProgramsDelay",bDef);
				if (bDef)
					SystemParametersInfo(SPI_GETMENUSHOWDELAY,NULL,&time,0);
				SetTimer(g_ProgramsButton,'CLSM',time,NULL);
			}
			g_LastHoverPos=pos;
		}
		if (msg->message==WM_TIMER && msg->wParam=='CLSM' && msg->hwnd==g_ProgramsButton)
		{
			g_bAllProgramsTimer=false;
			KillTimer(g_ProgramsButton,'CLSM');
			DWORD pos=GetMessagePos();
			if (pos==g_LastHoverPos)
				PostMessage(g_AllPrograms,WM_COMMAND,IDOK,(LPARAM)g_ProgramsButton);
			msg->message=WM_NULL;
		}
		if (msg->message==WM_MOUSELEAVE && msg->hwnd==g_ProgramsButton)
		{
			g_bAllProgramsTimer=false;
			KillTimer(g_ProgramsButton,'CLSM');
		}

		// handle hover
		if (msg->message==WM_MOUSEMOVE && msg->hwnd==g_StartButton && !g_bStartButtonTimer && !CMenuContainer::IsMenuOpened()
			&& GetSettingInt(L"Hover") && !WindowsMenuOpened())
		{
			g_bStartButtonTimer=true;
			int time=GetSettingInt(L"StartHoverDelay");
			SetTimer(g_StartButton,'CLSM',time,NULL);
		}
		if (msg->message==WM_MOUSELEAVE && msg->hwnd==g_StartButton)
		{
			g_bStartButtonTimer=false;
			KillTimer(g_StartButton,'CLSM');
		}
		if ((msg->message==WM_NCMOUSEMOVE || msg->message==WM_NCMOUSELEAVE) && msg->hwnd==g_TaskBar && GetSettingInt(L"Hover")
			&& (msg->wParam==HTCAPTION || !IsAppThemed())) // HACK: in Classic mode the start menu can show up even if wParam is not HTCAPTION (most likely a bug in Windows)
		{
			if (!CMenuContainer::IsMenuOpened() && !WindowsMenuOpened() && PointAroundStartButton())
			{
				if (!g_bStartButtonTimer)
				{
					g_bStartButtonTimer=true;
					int time=GetSettingInt(L"StartHoverDelay");
					SetTimer(g_StartButton,'CLSM',time,NULL);
				}
			}
			else
			{
				if (g_bStartButtonTimer)
				{
					g_bStartButtonTimer=false;
					KillTimer(g_StartButton,'CLSM');
				}
			}
		}
		if (msg->message==WM_TIMER && msg->wParam=='CLSM' && msg->hwnd==g_StartButton)
		{
			KillTimer(g_StartButton,'CLSM');
			msg->message=WM_NULL;
			if (g_bStartButtonTimer && !CMenuContainer::IsMenuOpened() && !WindowsMenuOpened())
			{
				CPoint pt(GetMessagePos());
				if (WindowFromPoint(pt)==g_StartButton || PointAroundStartButton())
				{
					int control=GetSettingInt(L"Hover");
					if (control==OPEN_CLASSIC && GetAsyncKeyState(VK_CONTROL)>=0)
					{
						// simulate Ctrl+Click. we can't simply show the menu here, because a window can't be activated without being clicked
						INPUT inputs[4]={
							{INPUT_KEYBOARD},
							{INPUT_MOUSE},
							{INPUT_MOUSE},
							{INPUT_KEYBOARD},
						};
						bool bSwap=GetSystemMetrics(SM_SWAPBUTTON)!=0;
						inputs[0].ki.wVk=VK_CONTROL;
						inputs[1].mi.dwFlags=bSwap?MOUSEEVENTF_RIGHTDOWN:MOUSEEVENTF_LEFTDOWN;
						inputs[2].mi.dwFlags=bSwap?MOUSEEVENTF_RIGHTUP:MOUSEEVENTF_LEFTUP;
						inputs[3].ki.wVk=VK_CONTROL;
						inputs[3].ki.dwFlags=KEYEVENTF_KEYUP;
						SendInput(_countof(inputs),inputs,sizeof(INPUT));
					}
					else if (control==OPEN_WINDOWS)
					{
						FindWindowsMenu();
						PostMessage(g_ProgWin,WM_SYSCOMMAND,SC_TASKLIST,'CLSM');
					}
				}
			}
			g_bStartButtonTimer=false;
		}

		if (msg->message==WM_NCRBUTTONUP && g_SwitchList && IsWindow(g_SwitchList))
		{
			wchar_t className[256]={0};
			GetClassName(g_SwitchList,className,_countof(className));
			if (wcscmp(className,L"ImmersiveSwitchList")==0)
			{
				msg->message=WM_NULL;
				ShowWindow(g_SwitchList,SW_SHOW);
				CPoint pt(GetMessagePos());
				ScreenToClient(g_SwitchList,&pt);
				PostMessage(g_SwitchList,WM_RBUTTONUP,wParam,MAKELONG(pt.x,pt.y));
			}
		}
		if (msg->message==WM_RBUTTONUP && msg->hwnd==g_StartButton)
		{
			WPARAM bShift=(GetSettingInt(L"MouseClick")!=OPEN_CLASSIC)?MK_SHIFT:0;
			if ((msg->wParam&MK_SHIFT)==bShift)
			{
				// additional commands for the context menu
				enum
				{
					CMD_SETTINGS=1,
					CMD_HELP,
					CMD_EXIT,
					CMD_OPEN,
					CMD_OPEN_ALL,
					CMD_EXPLORER,
				};

				// right-click on the start button - open the context menu (Settings, Help, Exit)
				msg->message=WM_NULL;
				POINT p={(short)LOWORD(msg->lParam),(short)HIWORD(msg->lParam)};
				ClientToScreen(msg->hwnd,&p);
				HMENU menu=CreatePopupMenu();
				CString title=LoadStringEx(IDS_MENU_TITLE);
				if (!title.IsEmpty())
				{
					AppendMenu(menu,MF_STRING,0,title);
					EnableMenuItem(menu,0,MF_BYPOSITION|MF_DISABLED);
					SetMenuDefaultItem(menu,0,TRUE);
					AppendMenu(menu,MF_SEPARATOR,0,0);
				}
				int count0=GetMenuItemCount(menu);
				if (GetSettingBool(L"EnableExplorer"))
				{
					if (!GetSettingString(L"ExplorerPath").IsEmpty())
						AppendMenu(menu,MF_STRING,CMD_EXPLORER,FindTranslation(L"Menu.Explorer",L"Windows Explorer"));
					AppendMenu(menu,MF_STRING,CMD_OPEN,FindTranslation(L"Menu.Open",L"&Open"));
					if (!SHRestricted(REST_NOCOMMONGROUPS))
						AppendMenu(menu,MF_STRING,CMD_OPEN_ALL,FindTranslation(L"Menu.OpenAll",L"O&pen All Users"));
					AppendMenu(menu,MF_SEPARATOR,0,0);
				}
				if (GetSettingBool(L"EnableSettings"))
					AppendMenu(menu,MF_STRING,CMD_SETTINGS,FindTranslation(L"Menu.MenuSettings",L"Settings"));
				if (HasHelp())
					AppendMenu(menu,MF_STRING,CMD_HELP,FindTranslation(L"Menu.MenuHelp",L"Help"));
				if (GetSettingBool(L"EnableExit"))
					AppendMenu(menu,MF_STRING,CMD_EXIT,FindTranslation(L"Menu.MenuExit",L"Exit"));
				if (GetMenuItemCount(menu)>count0)
				{
					MENUITEMINFO mii={sizeof(mii)};
					mii.fMask=MIIM_BITMAP;
					mii.hbmpItem=HBMMENU_POPUP_CLOSE;
					SetMenuItemInfo(menu,CMD_EXIT,FALSE,&mii);
					MENUINFO info={sizeof(info),MIM_STYLE,MNS_CHECKORBMP};
					SetMenuInfo(menu,&info);
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
						ShowHelp();
						return TRUE;
					}
					if (res==CMD_EXIT)
					{
						LRESULT res=CallNextHookEx(NULL,code,wParam,lParam);
						CleanStartMenuDLL();
						return res; // we should exit as quickly as possible now. the DLL is about to be unloaded
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
					if (res==CMD_EXPLORER)
					{
						CString path=GetSettingString(L"ExplorerPath");
						ITEMIDLIST blank={0};
						SHELLEXECUTEINFO execute={sizeof(execute)};
						execute.lpVerb=L"open";
						execute.lpFile=path;
						execute.nShow=SW_SHOWNORMAL;
						if (_wcsicmp(path,L"computer")==0)
							execute.lpFile=L"::{20D04FE0-3AEA-1069-A2D8-08002B30309D}";
						else if (_wcsicmp(path,L"libraries")==0)
							execute.lpFile=L"::{031E4825-7B94-4DC3-B131-E946B44C8DD5}";
						else if (_wcsicmp(path,L"desktop")==0)
						{
							execute.fMask=SEE_MASK_IDLIST;
							execute.lpIDList=&blank;
							execute.lpFile=NULL;
						}
						else
						{
							execute.fMask=SEE_MASK_DOENVSUBST;
						}
						ShellExecuteEx(&execute);
					}
				}
			}
			else if (g_StartButtonOld)
			{
				CPoint pt(GetMessagePos());
				ScreenToClient(g_StartButtonOld,&pt);
				PostMessage(g_StartButtonOld,WM_RBUTTONUP,wParam,MAKELONG(pt.x,pt.y));
			}
		}

	}
	return CallNextHookEx(NULL,code,wParam,lParam);
}

// WH_GETMESSAGE hook for the explorer's GUI thread. The start menu exe uses this hook to inject code into the explorer process
STARTMENUAPI LRESULT CALLBACK HookInject( int code, WPARAM wParam, LPARAM lParam )
{
	if (code==HC_ACTION && !g_TaskBar)
		InitStartMenuDLL();
	return CallNextHookEx(NULL,code,wParam,lParam);
}
