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

static HWND HookStartMenu( bool bHookExplorer )
{
	HMODULE hHookModule=GetModuleHandle(L"ClassicStartMenuDLL.dll");

	// find the Progman window and the start button
	HWND progWin=FindWindowEx(NULL,NULL,L"Progman",NULL);
	if (!progWin) return NULL; // the Progman window may not be created yet (if Explorer is currently restarting)

	DWORD process;
	DWORD thread=GetWindowThreadProcessId(progWin,&process);

	if (!FindTaskBar(process)) return NULL; // the taskbar may not be created yet (if Explorer is currently restarting)

	if (!bHookExplorer)
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
		HookStartMenu(true);
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
/*		STARTUPINFO startupInfo={sizeof(STARTUPINFO)};
		PROCESS_INFORMATION processInfo;
		FILE *f=fopen("D:\\start.txt","wt");
		if(CreateProcess(L"C:\\windows\\explorer.exe",L"C:\\windows\\explorer.exe shell:::{3080F90D-D7AD-11D9-BD98-0000947B0257}",NULL,NULL,TRUE,NORMAL_PRIORITY_CLASS,NULL,NULL,&startupInfo,&processInfo))
		{
			fprintf(f,"success\n");
			CloseHandle(processInfo.hProcess);
			CloseHandle(processInfo.hThread);
		}
		else
		{
			int err=GetLastError();
			fprintf(f,"error %d\n",err);
		}
		fclose(f);
		CoInitialize(NULL);
		ShellExecute(NULL,NULL,L"explorer.exe",L"shell:::{3080F90D-D7AD-11D9-BD98-0000947B0257}",NULL,SW_SHOWNORMAL);
		CoUninitialize();*/
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
		if ((DWORD_PTR)ShellExecute(NULL,NULL,exe,args,NULL,SW_SHOWNORMAL)>32 && !args && LOWORD(GetVersion())!=0x0006)
		{
			// on Windows 7 the executed documents are not automatically added to the recent document list
			CComPtr<IShellLink> pLink;
			if (SUCCEEDED(CoCreateInstance(CLSID_ShellLink,NULL,CLSCTX_INPROC_SERVER,IID_IShellLink,(LPVOID*)&pLink)))
			{
				// assume it is a link and try to get the target path
				CComQIPtr<IPersistFile> pFile=pLink;
				if (pFile && SUCCEEDED(pFile->Load(exe,STGM_READ)))
					pLink->GetPath(exe,_countof(exe),NULL,0);
			}
			SHAddToRecentDocs(SHARD_PATH,exe);
		}

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
		EditSettings(true);
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
		ChangeWindowMessageFilterEx(window,WM_CLOSE,MSGFLT_ADD,NULL);
	}
	else
	{
		ChangeWindowMessageFilter(g_TaskbarCreatedMsg,MSGFLT_ADD);
		ChangeWindowMessageFilter(WM_CLEAR,MSGFLT_ADD);
		ChangeWindowMessageFilter(WM_CLOSE,MSGFLT_ADD);
	}

	MSG msg;
	HWND menu=HookStartMenu(bHookExplorer);
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
