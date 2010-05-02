// Classic Shell (c) 2009-2010, Ivo Beltchev
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

#include "ClassicStartMenuDLL\ClassicStartMenuDLL.h"

static HHOOK g_StartHook;
static HWND g_StartButton;

// MiniDumpNormal - minimal information
// MiniDumpWithDataSegs - include global variables
// MiniDumpWithFullMemory - include heap
static MINIDUMP_TYPE MiniDumpType=MiniDumpNormal;

static DWORD WINAPI SaveCrashDump( void *pExceptionInfo )
{
	HMODULE dbghelp=NULL;
	{
		wchar_t path[_MAX_PATH];

		if (GetModuleFileName(NULL,path,_MAX_PATH))
		{
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
					wchar_t fname[_MAX_FNAME];
					Sprintf(fname,_countof(fname),L"CSM_Crash%d.dmp",i);
					*PathFindFileName(path)=0;
					Strcat(path,_countof(path),fname);
					file=CreateFile(path,GENERIC_WRITE,FILE_SHARE_WRITE,NULL,CREATE_NEW,FILE_ATTRIBUTE_NORMAL,NULL);
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
	}
	if (dbghelp) FreeLibrary(dbghelp);
	TerminateProcess(GetCurrentProcess(),10);
	return 0;
}

static LONG _stdcall TopLevelFilter( _EXCEPTION_POINTERS *pExceptionInfo )
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

static void UnhookStartMenu( void )
{
	if (g_StartHook)
		UnhookWindowsHookEx(g_StartHook);
	g_StartHook=NULL;
}

static HWND HookStartMenu( bool bHookExplorer )
{
	HMODULE hHookModule=GetModuleHandle(L"ClassicStartMenuDLL.dll");

	// find the Progman window and the start button
	HWND progWin=FindWindowEx(NULL,NULL,L"Progman",NULL);
	if (!progWin) return NULL; // the Progman window may not be created yet (if Explorer is currently restarting)

	DWORD process;
	DWORD thread=GetWindowThreadProcessId(progWin,&process);

	g_StartButton=FindStartButton(process);
	if (!g_StartButton) return NULL; // the start button may not be created yet (if Explorer is currently restarting)

	if (!bHookExplorer)
		return ToggleStartMenu(g_StartButton,false);

	// install hooks in the explorer process
	thread=GetWindowThreadProcessId(g_StartButton,NULL);
	g_StartHook=SetWindowsHookEx(WH_GETMESSAGE,HookInject,hHookModule,thread);
	PostMessage(g_StartButton,WM_NULL,0,0); // make sure there is one message in the queue

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
	if (g_StartButton) ::PostMessage(g_StartButton,RegisterWindowMessage(L"ClassicStartMenu.StartMenuMsg"),wParam,lParam);
	return 0;
}

LRESULT CStartHookWindow::OnClose( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled )
{
	UnhookStartMenu();
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
		HookStartMenu(true);
		if (g_StartHook)
			KillTimer(TIMER_HOOK);
	}
	return 0;
}

int WINAPI wWinMain( HINSTANCE hInstance, HINSTANCE hPrevInstance, LPTSTR lpstrCmdLine, int nCmdShow )
{
	int open=-1;
	if (wcsstr(lpstrCmdLine,L"-togglenew")!=NULL) open=-2;
	else if (wcsstr(lpstrCmdLine,L"-toggle")!=NULL) open=0;
	else if (wcsstr(lpstrCmdLine,L"-open")!=NULL) open=1;

	bool bHookExplorer=!wcsstr(lpstrCmdLine,L"-nohook");

	if (!bHookExplorer)
		SetUnhandledExceptionFilter(TopLevelFilter);

	// prevent multiple instances from hooking the same explorer process
	HWND progWin=FindWindowEx(NULL,NULL,L"Progman",NULL);
	DWORD process;
	DWORD thread=GetWindowThreadProcessId(progWin,&process);

	if (bHookExplorer)
	{
		wchar_t mutexName[256];
		Sprintf(mutexName,_countof(mutexName),L"ClassicStartMenu.Mutex.%08x",process);
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
	}

	OleInitialize(NULL);
	g_TaskbarCreatedMsg=RegisterWindowMessage(L"TaskbarCreated");
	ChangeWindowMessageFilter(g_TaskbarCreatedMsg,MSGFLT_ADD);
	CStartHookWindow window;
	window.Create(NULL,NULL,L"StartHookWindow",WS_POPUP);

	MSG msg;
	HWND menu=HookStartMenu(bHookExplorer);
	if (bHookExplorer && open>=0)
		window.PostMessage(WM_OPEN,open,0);
	while ((bHookExplorer || IsWindow(menu)) && GetMessage(&msg,0,0,0))
	{
		TranslateMessage(&msg);
		DispatchMessage(&msg);
	}
	window.DestroyWindow();
	OleUninitialize();
	return 0;
}
