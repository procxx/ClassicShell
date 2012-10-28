// Classic Shell (c) 2009-2012, Ivo Beltchev
// The sources for Classic Shell are distributed under the MIT open source license

#include <windows.h>
#include <atlbase.h>
#include <atlcom.h>
#include <atlctl.h>
#include <atlstr.h>
#include <commctrl.h>
#include <shlobj.h>
#include <dbghelp.h>
#include "StringUtils.h"
#include "Settings.h"
#include "ResourceHelper.h"

#include "ClassicStartMenuDLL\ClassicStartMenuDLL.h"
#include "ClassicStartMenuDLL\SettingsUI.h"

static HHOOK g_StartHook;

static void UnhookStartMenu( void )
{
	if (g_StartHook)
		UnhookWindowsHookEx(g_StartHook);
	g_StartHook=NULL;
}

enum THookMode
{
	HOOK_NONE, // don't hook Explorer, running as a separate exe
	HOOK_NORMAL, // hook Explorer normally, no retry
	HOOK_STARTUP, // retry to hook Explorer
};

static HWND HookStartMenu( THookMode mode )
{
	HMODULE hHookModule=GetModuleHandle(L"ClassicStartMenuDLL.dll");

	// find the Progman window and the start button

	HWND progWin;
	bool bFindAppManager=(mode==HOOK_STARTUP && GetWinVersion()>=WIN_VER_WIN8);
	for (int i=0;i<120;i++) // retry for 1 minute
	{
		if (bFindAppManager)
			bFindAppManager=!FindWindow(L"ApplicationManager_DesktopShellWindow",NULL);
		if (!bFindAppManager)
		{
			progWin=FindWindowEx(NULL,NULL,L"Progman",NULL);
			if (progWin) break;
			if (mode!=HOOK_STARTUP) return NULL; // the Progman window may not be created yet (if Explorer is currently restarting)
		}
		Sleep(500);
	}
	DWORD process;
	DWORD thread=GetWindowThreadProcessId(progWin,&process);

	for (int i=0;i<10;i++) // retry for 5 sec
	{
		if (FindTaskBar(process)) break;
		if (mode!=HOOK_STARTUP) return NULL; // the taskbar may not be created yet (if Explorer is currently restarting)
		Sleep(500);
	}

	if (mode==HOOK_NONE)
		return ToggleStartMenu(g_StartButton,false);

	// install hooks in the explorer process
	thread=GetWindowThreadProcessId(g_TaskBar,NULL);
	g_StartHook=SetWindowsHookEx(WH_GETMESSAGE,HookInject,hHookModule,thread);
	if (!g_StartHook)
	{
		int err=GetLastError();
		LogHookError(err);
	}
	PostMessage(g_TaskBar,WM_NULL,0,0); // make sure there is one message in the queue

	return NULL;
}

static UINT g_TaskbarCreatedMsg; // the "TaskbarCreated" message

// CStartHookWindow is a hidden window that waits for the "TaskbarCreated" message and rehooks the explorer process
// Also when the start menu wants to shut down it sends WM_CLOSE to this window, which unhooks explorer and exits

const int WM_OPEN=WM_USER+10;

const int TIMER_HOOK=1;

class CStartHookWindow: public CWindowImpl<CStartHookWindow>
{
public:

	DECLARE_WND_CLASS(L"ClassicStartMenu.CStartHookWindow")

	BEGIN_MSG_MAP( CStartHookWindow )
		MESSAGE_HANDLER( WM_OPEN, OnOpen )
		MESSAGE_HANDLER( WM_CLOSE, OnClose )
		MESSAGE_HANDLER( WM_CLEAR, OnClear )
		MESSAGE_HANDLER( WM_TIMER, OnTimer )
		MESSAGE_HANDLER( g_TaskbarCreatedMsg, OnTaskbarCreated )
	END_MSG_MAP()

protected:
	// Handler prototypes:
	//  LRESULT MessageHandler(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
	//  LRESULT CommandHandler(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled);
	//  LRESULT NotifyHandler(int idCtrl, LPNMHDR pnmh, BOOL& bHandled);
	LRESULT OnOpen( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled );
	LRESULT OnClose( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled );
	LRESULT OnClear( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled );
	LRESULT OnTimer( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled );
	LRESULT OnTaskbarCreated( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled );
};

LRESULT CStartHookWindow::OnOpen( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled )
{
	if (g_TaskBar) ::PostMessage(g_TaskBar,RegisterWindowMessage(L"ClassicStartMenu.StartMenuMsg"),wParam,lParam);
	return 0;
}

LRESULT CStartHookWindow::OnClose( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled )
{
	UnhookStartMenu();
	Sleep(100);
	PostQuitMessage(0);
	return 0;
}

LRESULT CStartHookWindow::OnClear( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled )
{
	UnhookStartMenu();
	return 0;
}

LRESULT CStartHookWindow::OnTaskbarCreated( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled )
{
	SetTimer(TIMER_HOOK,100);
	return 0;
}

LRESULT CStartHookWindow::OnTimer( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled )
{
	if (wParam==TIMER_HOOK)
	{
		UnhookStartMenu();
		HookStartMenu(HOOK_NORMAL);
		if (g_StartHook)
			KillTimer(TIMER_HOOK);
	}
	return 0;
}

enum
{
	CMD_NONE=-1,
	CMD_TOGGLE_NEW=-2,
};

int WINAPI wWinMain( HINSTANCE hInstance, HINSTANCE hPrevInstance, LPTSTR lpstrCmdLine, int nCmdShow )
{
	if (wcsstr(lpstrCmdLine,L"-startup"))
	{
	}
	wchar_t path[_MAX_PATH];
	GetModuleFileName(NULL,path,_countof(path));
	*PathFindFileName(path)=0;
	SetCurrentDirectory(path);
	const wchar_t *pRunAs=wcsstr(lpstrCmdLine,L"-runas");
	if (pRunAs)
	{
		CoInitialize(NULL);
		wchar_t exe[_MAX_PATH];
		const wchar_t *args=SeparateArguments(pRunAs+7,exe);
		SHELLEXECUTEINFO execute={sizeof(execute),SEE_MASK_DOENVSUBST|SEE_MASK_FLAG_LOG_USAGE};
		execute.lpFile=exe;
		execute.lpParameters=args;
		execute.nShow=SW_SHOWNORMAL;
		ShellExecuteEx(&execute);
		CoUninitialize();
		return 0;
	}
	int open=CMD_NONE;
	if (wcsstr(lpstrCmdLine,L"-togglenew")!=NULL) open=CMD_TOGGLE_NEW;
	else if (wcsstr(lpstrCmdLine,L"-toggle")!=NULL) open=MSG_TOGGLE;
	else if (wcsstr(lpstrCmdLine,L"-open")!=NULL) open=MSG_OPEN;
	else if (wcsstr(lpstrCmdLine,L"-settings")!=NULL) open=MSG_SETTINGS;
	else if (wcsstr(lpstrCmdLine,L"-exit")!=NULL) open=MSG_EXIT;

	const wchar_t *pNoHook=wcsstr(lpstrCmdLine,L"-nohook");
	bool bHookExplorer=!pNoHook;
	if (pNoHook)
	{
		pNoHook+=7;
		if (*pNoHook=='1') MiniDumpType=MiniDumpNormal;
		if (*pNoHook=='2') MiniDumpType=MiniDumpWithDataSegs;
		if (*pNoHook=='3') MiniDumpType=MiniDumpWithFullMemory;
	}

	if (!bHookExplorer)
		SetUnhandledExceptionFilter(TopLevelFilter);

#ifndef BUILD_SETUP
	if (wcsstr(lpstrCmdLine,L"-testsettings")!=NULL)
	{
		CoInitialize(NULL);
		EditSettings(true,0);
		CoUninitialize();
		return 0;
	}
#endif

	// prevent multiple instances from running on the same desktop
	// the assumption is that multiple desktops for the same user will have different name (but may repeat across users)
	wchar_t userName[256];
	DWORD len=_countof(userName);
	GetUserName(userName,&len);
	len=0;
	HANDLE desktop=GetThreadDesktop(GetCurrentThreadId());
	GetUserObjectInformation(desktop,UOI_NAME,NULL,0,&len);
	wchar_t *deskName=(wchar_t*)malloc(len);
	GetUserObjectInformation(desktop,UOI_NAME,deskName,len,&len);

	wchar_t mutexName[1024];
	Sprintf(mutexName,_countof(mutexName),L"ClassicStartMenu.Mutex.%s.%s",userName,deskName);
	free(deskName);

	HWND progWin=FindWindowEx(NULL,NULL,L"Progman",NULL);
	DWORD process;
	if (progWin)
		GetWindowThreadProcessId(progWin,&process);

	HANDLE hMutex=NULL;
	if (bHookExplorer)
	{
		hMutex=CreateMutex(NULL,TRUE,mutexName);
		if (GetLastError()==ERROR_ALREADY_EXISTS || GetLastError()==ERROR_ACCESS_DENIED)
		{
			if (open==CMD_TOGGLE_NEW)
			{
				if (progWin)
				{
					AllowSetForegroundWindow(process);
					PostMessage(progWin,WM_SYSCOMMAND,SC_TASKLIST,'CLSM');
				}
			}
			else if (open!=CMD_NONE)
			{
				AllowSetForegroundWindow(process);
				HWND hwnd=FindWindow(L"ClassicStartMenu.CStartHookWindow",L"StartHookWindow");
				if (hwnd) PostMessage(hwnd,WM_OPEN,open,0);
			}
			if (open==MSG_EXIT && hMutex && WaitForSingleObject(hMutex,2000)==WAIT_OBJECT_0)
				ReleaseMutex(hMutex);
			return 0;
		}
	}
	if (open!=CMD_NONE && open!=MSG_OPEN && open!=MSG_SETTINGS)
	{
		if (hMutex) ReleaseMutex(hMutex);
		return 0;
	}

	OleInitialize(NULL);
	CStartHookWindow window;
	window.Create(NULL,NULL,L"StartHookWindow",WS_POPUP);

	g_TaskbarCreatedMsg=RegisterWindowMessage(L"TaskbarCreated");
	typedef BOOL (WINAPI *tChangeWindowMessageFilterEx)(HWND hWnd, UINT message, DWORD action, PCHANGEFILTERSTRUCT pChangeFilterStruct );
	HMODULE hUser32=GetModuleHandle(L"user32.dll");
	tChangeWindowMessageFilterEx ChangeWindowMessageFilterEx=(tChangeWindowMessageFilterEx)GetProcAddress(hUser32,"ChangeWindowMessageFilterEx");
	if (ChangeWindowMessageFilterEx)
	{
		ChangeWindowMessageFilterEx(window,g_TaskbarCreatedMsg,MSGFLT_ADD,NULL);
		ChangeWindowMessageFilterEx(window,WM_CLEAR,MSGFLT_ADD,NULL);
		ChangeWindowMessageFilterEx(window,WM_OPEN,MSGFLT_ADD,NULL);
		ChangeWindowMessageFilterEx(window,WM_CLOSE,MSGFLT_ADD,NULL);
	}
	else
	{
		ChangeWindowMessageFilter(g_TaskbarCreatedMsg,MSGFLT_ADD);
		ChangeWindowMessageFilter(WM_CLEAR,MSGFLT_ADD);
		ChangeWindowMessageFilter(WM_OPEN,MSGFLT_ADD);
		ChangeWindowMessageFilter(WM_CLOSE,MSGFLT_ADD);
	}

	MSG msg;
	HWND menu=HookStartMenu(bHookExplorer?HOOK_STARTUP:HOOK_NONE);
	if (bHookExplorer && open>=0)
		window.PostMessage(WM_OPEN,open,MSG_OPEN);
	while ((bHookExplorer || IsWindow(menu)) && GetMessage(&msg,0,0,0))
	{
		TranslateMessage(&msg);
		DispatchMessage(&msg);
	}
	window.DestroyWindow();
	OleUninitialize();

	if (hMutex) ReleaseMutex(hMutex);
	return 0;
}
