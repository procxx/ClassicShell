// Classic Shell (c) 2009-2010, Ivo Beltchev
// The sources for Classic Shell are distributed under the MIT open source license

#include <windows.h>
#include <atlbase.h>
#include <atlcom.h>
#include <atlctl.h>
#include <atlstr.h>
#include <commctrl.h>
#include <shlobj.h>

#include "ClassicStartMenuDLL\ClassicStartMenuDLL.h"

#define HOOK_EXPLORER // when this is not defined the start menu runs directly in this process (for debugging)

#if defined(BUILD_SETUP) && !defined(HOOK_EXPLORER)
#define HOOK_EXPLORER // make sure it is defined in Setup
#endif

static HHOOK g_ProgHook, g_StartHook;
static HWND g_StartButton;

static void UnhookStartMenu( void )
{
	if (g_StartHook)
		UnhookWindowsHookEx(g_StartHook);
	g_StartHook=NULL;
	if (g_ProgHook)
		UnhookWindowsHookEx(g_ProgHook);
	g_ProgHook=NULL;
}

static HWND HookStartMenu( void )
{
	HMODULE hHookModule=GetModuleHandle(L"ClassicStartMenuDLL.dll");

	// find the Progman window and the start button
	HWND progWin=FindWindowEx(NULL,NULL,L"Progman",NULL);

	DWORD process;
	DWORD thread=GetWindowThreadProcessId(progWin,&process);

	g_StartButton=FindStartButton(process);

#ifdef HOOK_EXPLORER
	// install hooks in the explorer process
	g_ProgHook=SetWindowsHookEx(WH_GETMESSAGE,HookProgMan,hHookModule,thread);

	thread=GetWindowThreadProcessId(g_StartButton,NULL);
	g_StartHook=SetWindowsHookEx(WH_GETMESSAGE,HookStartButton,hHookModule,thread);

	return NULL;
#else
	return ToggleStartMenu(g_StartButton,false);
#endif
}

static UINT g_TaskbarCreatedMsg; // the "TaskbarCreated" message

// CStartHookWindow is a hidden window that waits for the "TaskbarCreated" message and rehooks the explorer process
// Also when the start menu wants to shut down it sends WM_CLOSE to this window, which unhooks explorer and exits

const int WM_OPEN=WM_USER+10;

class CStartHookWindow: public CWindowImpl<CStartHookWindow>
{
public:

	DECLARE_WND_CLASS(L"ClassicStartMenu.CStartHookWindow")

	BEGIN_MSG_MAP( CStartHookWindow )
		MESSAGE_HANDLER( WM_OPEN, OnOpen )
		MESSAGE_HANDLER( WM_CLOSE, OnClose )
		MESSAGE_HANDLER( g_TaskbarCreatedMsg, OnTaskbarCreated )
	END_MSG_MAP()

protected:
	// Handler prototypes:
	//  LRESULT MessageHandler(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
	//  LRESULT CommandHandler(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled);
	//  LRESULT NotifyHandler(int idCtrl, LPNMHDR pnmh, BOOL& bHandled);
	LRESULT OnOpen( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled );
	LRESULT OnClose( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled );
	LRESULT OnTaskbarCreated( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled );
};

LRESULT CStartHookWindow::OnOpen( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled )
{
	if (g_StartButton) ::PostMessage(g_StartButton,RegisterWindowMessage(L"ClassicStartMenu.StartMenuMsg"),wParam,lParam);
	return 0;
}

LRESULT CStartHookWindow::OnClose( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled )
{
	UnhookStartMenu();
	PostQuitMessage(0);
	return 0;
}

LRESULT CStartHookWindow::OnTaskbarCreated( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled )
{
	UnhookStartMenu();
	HookStartMenu();
	return 0;
}

int WINAPI wWinMain( HINSTANCE hInstance, HINSTANCE hPrevInstance, LPTSTR lpstrCmdLine, int nCmdShow )
{
	int open=-1;
	if (wcsstr(lpstrCmdLine,L"-togglenew")!=NULL) open=-2;
	else if (wcsstr(lpstrCmdLine,L"-toggle")!=NULL) open=0;
	else if (wcsstr(lpstrCmdLine,L"-open")!=NULL) open=1;
	// prevent multiple instances from hooking the same explorer process
	HWND progWin=FindWindowEx(NULL,NULL,L"Progman",NULL);
	DWORD process;
	DWORD thread=GetWindowThreadProcessId(progWin,&process);
	wchar_t mutexName[256];
	swprintf_s(mutexName,L"ClassicStartMenu.Mutex.%08x",process);
	HANDLE hMutex=CreateMutex(NULL,FALSE,mutexName);
	if (GetLastError()==ERROR_ALREADY_EXISTS || GetLastError()==ERROR_ACCESS_DENIED)
	{
		if (open>=0)
		{
			AllowSetForegroundWindow(process);
			HWND hwnd=FindWindow(L"ClassicStartMenu.CStartHookWindow",L"StartHookWindow");
			if (hwnd) PostMessage(hwnd,WM_OPEN,open,0);
		}
		if (open==-2 && progWin)
		{
			AllowSetForegroundWindow(process);
			PostMessage(progWin,WM_SYSCOMMAND,SC_TASKLIST,'CLSM');
		}
		return 0;
	}

	OleInitialize(NULL);
	g_TaskbarCreatedMsg=RegisterWindowMessage(L"TaskbarCreated");
	ChangeWindowMessageFilter(g_TaskbarCreatedMsg,MSGFLT_ADD);
	CStartHookWindow window;
	window.Create(NULL,NULL,L"StartHookWindow",WS_POPUP);

	MSG msg;
#ifdef HOOK_EXPLORER
	window.PostMessage(g_TaskbarCreatedMsg,0,0);
	if (open>=0) window.PostMessage(WM_OPEN,open,0);
	while (GetMessage(&msg,0,0,0))
#else
	HWND menu=HookStartMenu();
	while (IsWindow(menu) && GetMessage(&msg,0,0,0))
#endif
	{
		TranslateMessage(&msg);
		DispatchMessage(&msg);
	}
	window.DestroyWindow();
	OleUninitialize();
	return 0;
}
