// Classic Shell (c) 2009-2010, Ivo Beltchev
// The sources for Classic Shell are distributed under the MIT open source license

#include "stdafx.h"
#include "resource.h"
#include "ClassicStartMenuDLL.h"
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

#if defined(BUILD_SETUP) && !defined(HOOK_DROPTARGET)
#define HOOK_DROPTARGET // make sure it is defined in Setup
#endif

HWND g_StartButton;
HWND g_TaskBar;
HWND g_OwnerWindow;
HWND g_ProgWin;
HWND g_TopMenu, g_AllPrograms, g_ProgramsButton, g_UserPic; // from the Windows menu
static UINT g_StartMenuMsg;
static HWND g_Tooltip;
static TOOLINFO g_StarButtonTool;
static int g_HotkeyCSM, g_HotkeyWSM, g_HotkeyCSMID, g_HotkeyWSMID;
static HHOOK g_ProgHook, g_StartHook, g_KeyboardHook;
static bool g_bAllProgramsTimer;
static bool g_bStartButtonTimer;
static bool g_bInMenu;
static DWORD g_LastClickTime;
static DWORD g_LastHoverPos;
static bool g_bCrashDump;

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

DWORD WINAPI SaveCrashDump( void *pExceptionInfo )
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
		if (GetWindowThreadProcessId(g_TaskBar,NULL)==GetCurrentThreadId())
		{
			// find tooltip
			EnumThreadWindows(GetWindowThreadProcessId(g_TaskBar,NULL),FindTooltipEnum,NULL);
			if (g_Tooltip)
			{
				g_StarButtonTool.cbSize=sizeof(g_StarButtonTool);
				g_StarButtonTool.hwnd=g_TaskBar;
				g_StarButtonTool.uId=(UINT_PTR)g_StartButton;
				SendMessage(g_Tooltip,TTM_GETTOOLINFO,0,(LPARAM)&g_StarButtonTool);
			}
			g_OwnerWindow=g_Owner.Create(NULL,0,0,WS_POPUP,WS_EX_TOOLWINDOW|WS_EX_TOPMOST);
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
		PostMessage(g_StartButton,g_StartMenuMsg,(grfKeyState&MK_SHIFT)?MSG_SHIFTDRAG:MSG_DRAG,0);
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
		if (g_StartButton)
		{
#ifdef HOOK_DROPTARGET
			g_pOriginalTarget=(IDropTarget*)GetProp(g_StartButton,L"OleDropTargetInterface");
			if (g_pOriginalTarget)
				RevokeDragDrop(g_StartButton);
			CStartMenuTarget *pNewTarget=new CStartMenuTarget();
			RegisterDragDrop(g_StartButton,pNewTarget);
			pNewTarget->Release();
			g_HotkeyCSMID=GlobalAddAtom(L"ClassicStartMenu.HotkeyCSM");
			g_HotkeyWSMID=GlobalAddAtom(L"ClassicStartMenu.HotkeyWSM");
#endif
			EnableHotkeys(HOTKEYS_NORMAL);
			srand(GetTickCount());
		}
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
	return CMenuContainer::ToggleStartMenu(startButton,bKeyboard,false);
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
			PostMessage(g_StartButton,g_StartMenuMsg,MSG_SHIFTWIN,0);
		}
	}
	return CallNextHookEx(NULL,code,wParam,lParam);
}

// Set the hotkeys and controls for the start menu
void EnableHotkeys( THotkeys enable )
{
	if (!g_StartButton)
		return;
	if (GetWindowThreadProcessId(g_StartButton,NULL)!=GetCurrentThreadId())
	{
		PostMessage(g_StartButton,g_StartMenuMsg,MSG_HOTKEYS,enable);
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
		UnregisterHotKey(g_StartButton,g_HotkeyCSMID);
	g_HotkeyCSM=0;

	if (g_HotkeyWSM)
		UnregisterHotKey(g_StartButton,g_HotkeyWSMID);
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
			BOOL err=RegisterHotKey(g_StartButton,g_HotkeyCSMID,mod,g_HotkeyCSM&255);
			int q=0;
		}

		g_HotkeyWSM=GetSettingInt(L"WSMHotkey");
		if (g_HotkeyWSM)
		{
			int mod=0;
			if (g_HotkeyWSM&(HOTKEYF_SHIFT<<8)) mod|=MOD_SHIFT;
			if (g_HotkeyWSM&(HOTKEYF_CONTROL<<8)) mod|=MOD_CONTROL;
			if (g_HotkeyWSM&(HOTKEYF_ALT<<8)) mod|=MOD_ALT;
			RegisterHotKey(g_StartButton,g_HotkeyWSMID,mod,g_HotkeyWSM&255);
		}
	}
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
		}
		g_bAllProgramsTimer=false;
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

static  LRESULT CALLBACK SubclassProgramsProc( HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData )
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
	FindStartButton();
	g_ProgWin=FindWindowEx(NULL,NULL,L"Progman",NULL);
	DWORD thread=GetWindowThreadProcessId(g_ProgWin,NULL);
	g_ProgHook=SetWindowsHookEx(WH_GETMESSAGE,HookProgMan,NULL,thread);
	g_StartHook=SetWindowsHookEx(WH_GETMESSAGE,HookStartButton,NULL,GetCurrentThreadId());
	HWND hwnd=FindWindow(L"ClassicStartMenu.CStartHookWindow",L"StartHookWindow");
	LoadLibrary(L"ClassicStartMenuDLL.dll"); // keep the DLL from unloading
	if (hwnd) PostMessage(hwnd,WM_CLEAR,0,0); // tell the exe to unhook this hook
	SetWindowSubclass(g_TaskBar,SubclassTaskBarProc,'CLSH',0);
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
	g_IconManager.StopPreloading(true);
	UnhookDropTarget();
	EnableHotkeys(HOTKEYS_CLEAR);
	HWND hwnd=FindWindow(L"ClassicStartMenu.CStartHookWindow",L"StartHookWindow");
	UnhookWindowsHookEx(g_ProgHook);
	UnhookWindowsHookEx(g_StartHook);
	RemoveWindowSubclass(g_TaskBar,SubclassTaskBarProc,'CLSH');
	if (g_TopMenu)
	{
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
	g_TopMenu=NULL;

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
			FindStartButton();
			int control=GetSettingInt(L"WinKey");
			if (control!=OPEN_WINDOWS)
			{
				msg->message=WM_NULL;
				if (control==OPEN_CLASSIC)
					PostMessage(g_StartButton,g_StartMenuMsg,MSG_TOGGLE,0);
			}
			else
				PostMessage(g_StartButton,g_StartMenuMsg,MSG_FINDMENU,0);
		}
	}
	return CallNextHookEx(NULL,code,wParam,lParam);
}

// Returns true if the mouse is on the taskbar portion of the start button
static bool PointAroundStartButton( void )
{
	DWORD pos=GetMessagePos();
	POINT pt={(short)LOWORD(pos),(short)HIWORD(pos)};
	RECT rc;
	GetWindowRect(g_TaskBar,&rc);
	if (!PtInRect(&rc,pt))
		return false;

	// in Classic mode there are few pixels at the edge of the taskbar that are not covered by a child window
	// we nudge the point to be in the middle of the taskbar to avoid those pixels
	// also ignore clicks on the half of the taskbar that doesn't contain the start button
	APPBARDATA appbar={sizeof(appbar),g_TaskBar};
	SHAppBarMessage(ABM_GETTASKBARPOS,&appbar);
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
			if (!IsLanguageRTL())
				return false;
		}
		else
		{
			if (IsLanguageRTL())
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
		if (((msg->message>=WM_MOUSEFIRST && msg->message<=WM_MOUSELAST) || msg->message==WM_MOUSEHOVER || msg->message==WM_MOUSELEAVE) && CMenuContainer::ProcessMouseMessage(msg->hwnd,msg->message,msg->wParam,msg->lParam))
		{
			msg->message=WM_NULL;
			return 0;
		}
		if (msg->message==g_StartMenuMsg && msg->hwnd==g_StartButton)
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

		if (msg->message==WM_HOTKEY && msg->hwnd==g_StartButton)
		{
			if (msg->wParam==g_HotkeyCSMID)
			{
				msg->message=WM_NULL;
				SetForegroundWindow(g_StartButton);
				ToggleStartMenu(g_StartButton,true);
			}
			else if (msg->wParam==g_HotkeyWSMID)
				PostMessage(g_ProgWin,WM_SYSCOMMAND,SC_TASKLIST,'CLSM');
		}

		if (msg->message==WM_KEYDOWN && msg->hwnd==g_StartButton && (msg->wParam==VK_SPACE || msg->wParam==VK_RETURN))
		{
			FindWindowsMenu();
			if (GetSettingInt(L"WinKey")==OPEN_CLASSIC)
			{
				msg->message=WM_NULL;
				SetForegroundWindow(g_StartButton);
				ToggleStartMenu(g_StartButton,true);
			}
		}

		if ((msg->message==WM_LBUTTONDOWN || msg->message==WM_LBUTTONDBLCLK || msg->message==WM_MBUTTONDOWN) && msg->hwnd==g_StartButton)
		{
			FindWindowsMenu();
			const wchar_t *name;
			if (msg->message==WM_MBUTTONDOWN)
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
			else if (control==OPEN_WINDOWS && msg->message==WM_MBUTTONDOWN)
				PostMessage(g_ProgWin,WM_SYSCOMMAND,SC_TASKLIST,'CLSM');
		}

		if ((msg->message==WM_NCLBUTTONDOWN || msg->message==WM_NCLBUTTONDBLCLK || msg->message==WM_NCMBUTTONDOWN) && msg->hwnd==g_TaskBar
			&& (msg->wParam==HTCAPTION || !IsAppThemed())) // HACK: in Classic mode the start menu can show up even if wParam is not HTCAPTION (most likely a bug in Windows)
		{
			FindWindowsMenu();
			const wchar_t *name;
			if (msg->message==WM_MBUTTONDOWN)
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
			else if (control==OPEN_WINDOWS && msg->message==WM_NCMBUTTONDOWN)
				PostMessage(g_ProgWin,WM_SYSCOMMAND,SC_TASKLIST,'CLSM');
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
			&& GetSettingInt(L"Hover") && !(::SendMessage(g_StartButton,BM_GETSTATE,0,0)&BST_PUSHED))
		{
			g_bStartButtonTimer=true;
			int time=GetSettingInt(L"StartHoverDelay");
			SetTimer(g_StartButton,'CLSM',time,NULL);
		}
		if ((msg->message==WM_NCMOUSEMOVE || msg->message==WM_NCMOUSELEAVE) && msg->hwnd==g_TaskBar && GetSettingInt(L"Hover")
			&& (msg->wParam==HTCAPTION || !IsAppThemed())) // HACK: in Classic mode the start menu can show up even if wParam is not HTCAPTION (most likely a bug in Windows)
		{
			if (!CMenuContainer::IsMenuOpened() && !(::SendMessage(g_StartButton,BM_GETSTATE,0,0)&BST_PUSHED) && PointAroundStartButton())
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
			if (g_bStartButtonTimer && !CMenuContainer::IsMenuOpened() && !(::SendMessage(g_StartButton,BM_GETSTATE,0,0)&BST_PUSHED))
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
						inputs[0].ki.wVk=VK_CONTROL;
						inputs[1].mi.dwFlags=MOUSEEVENTF_LEFTDOWN;
						inputs[2].mi.dwFlags=MOUSEEVENTF_LEFTUP;
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
				};

				// right-click on the start button - open the context menu (Settings, Help, Exit)
				msg->message=WM_NULL;
				POINT p={(short)LOWORD(msg->lParam),(short)HIWORD(msg->lParam)};
				ClientToScreen(msg->hwnd,&p);
				HMENU menu=CreatePopupMenu();
				AppendMenu(menu,MF_STRING,0,LoadStringEx(IDS_MENU_TITLE));
				EnableMenuItem(menu,0,MF_BYPOSITION|MF_DISABLED);
				SetMenuDefaultItem(menu,0,TRUE);
				AppendMenu(menu,MF_SEPARATOR,0,0);
				AppendMenu(menu,MF_STRING,CMD_OPEN,FindTranslation(L"Menu.Open",L"&Open"));
				AppendMenu(menu,MF_STRING,CMD_OPEN_ALL,FindTranslation(L"Menu.OpenAll",L"O&pen All Users"));
				AppendMenu(menu,MF_SEPARATOR,0,0);
				if (GetSettingBool(L"EnableSettings"))
					AppendMenu(menu,MF_STRING,CMD_SETTINGS,FindTranslation(L"Menu.MenuSettings",L"Settings"));
				AppendMenu(menu,MF_STRING,CMD_HELP,FindTranslation(L"Menu.MenuHelp",L"Help"));
				if (GetSettingBool(L"EnableExit"))
					AppendMenu(menu,MF_STRING,CMD_EXIT,FindTranslation(L"Menu.MenuExit",L"Exit"));
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
					wchar_t path[_MAX_PATH];
					GetModuleFileName(g_Instance,path,_countof(path));
					*PathFindFileName(path)=0;
					Strcat(path,_countof(path),DOC_PATH L"ClassicShell.chm::/ClassicStartMenu.html");
					HtmlHelp(GetDesktopWindow(),path,HH_DISPLAY_TOPIC,NULL);
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
			}
		}

	}
	return CallNextHookEx(NULL,code,wParam,lParam);
}

// WH_GETMESSAGE hook for the explorer's GUI thread. The start menu exe uses this hook to inject code into the explorer process
STARTMENUAPI LRESULT CALLBACK HookInject( int code, WPARAM wParam, LPARAM lParam )
{
	if (code==HC_ACTION && !g_StartButton)
	{
		MSG *msg=(MSG*)lParam;
		InitStartMenuDLL();
	}
	return CallNextHookEx(NULL,code,wParam,lParam);
}
