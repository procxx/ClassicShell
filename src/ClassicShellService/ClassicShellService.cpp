#include <windows.h>
#include <WtsApi32.h>
#include <shlwapi.h>
#include <stdio.h>
#include <time.h>
#include <dbghelp.h>

static const wchar_t *g_ServiceName=L"ClassicShellService";
static SERVICE_STATUS_HANDLE g_hServiceStatus;
static SERVICE_STATUS g_ServiceStatus;
static wchar_t g_LogName[_MAX_PATH];

static void LogText( const char *format, ... )
{
	if (*g_LogName)
	{
		FILE *f;
		if (_wfopen_s(&f,g_LogName,L"a+t")) return;
		va_list args;
		va_start(args,format);
		fprintf(f,"0x%8X  ",time(NULL));
		vfprintf(f,format,args);
		va_end(args);
		fclose(f);
	}
}

static void StartStartMenu( DWORD sessionId )
{
	// run the classic start menu on logon
	HANDLE hUser;
	if (WTSQueryUserToken(sessionId,&hUser))
	{
		STARTUPINFO startupInfo={sizeof(STARTUPINFO),NULL,L"Winsta0\\Default"};
		PROCESS_INFORMATION processInfo;
		wchar_t path[_MAX_PATH];
		GetModuleFileName(NULL,path,_countof(path));
		PathRemoveFileSpec(path);
		PathAppend(path,L"ClassicStartMenu.exe -startup");
		LogText("Starting process: %S\n",path);
		if(CreateProcessAsUser(hUser,NULL,path,NULL,NULL,TRUE,NORMAL_PRIORITY_CLASS,NULL,NULL,&startupInfo,&processInfo))
		{
			CloseHandle(processInfo.hProcess);
			CloseHandle(processInfo.hThread);
		}
		else
		{
			int err=GetLastError();
			LogText("CreateProcessAsUser failed: %d\n",err);
		}
		CloseHandle(hUser);
	}
	else
	{
		int err=GetLastError();
		LogText("WTSQueryUserToken failed: %d\n",err);
	}
}

static DWORD WINAPI ServiceHandlerEx(DWORD dwControl, DWORD dwEventType, LPVOID lpEventData, LPVOID lpContext)
{
	LogText("Control: %d, Event: %d\n",dwControl,dwEventType);
	switch(dwControl)
	{
		case SERVICE_CONTROL_STOP:
		case SERVICE_CONTROL_SHUTDOWN:
			g_ServiceStatus.dwCurrentState=SERVICE_STOPPED;
			break;
		case SERVICE_CONTROL_PAUSE:
			g_ServiceStatus.dwCurrentState=SERVICE_PAUSED;
			break;
		case SERVICE_CONTROL_CONTINUE:
			g_ServiceStatus.dwCurrentState=SERVICE_RUNNING;
			break;
		case SERVICE_CONTROL_INTERROGATE:
			break;
		case SERVICE_CONTROL_SESSIONCHANGE:
			if (dwEventType==WTS_SESSION_LOGON)
			{
				StartStartMenu(((WTSSESSION_NOTIFICATION*)lpEventData)->dwSessionId);
			}
			return NO_ERROR;
		default:
			return ERROR_CALL_NOT_IMPLEMENTED;
	};
	SetServiceStatus(g_hServiceStatus,&g_ServiceStatus);
	return NO_ERROR;
}

static void WINAPI ServiceMain( DWORD dwArgc, LPTSTR *lpszArgv )
{
	WTS_SESSION_INFO *pInfos;
	DWORD count;
	if (WTSEnumerateSessions(WTS_CURRENT_SERVER_HANDLE,NULL,1,&pInfos,&count))
	{
		for (DWORD i=0;i<count;i++)
		{
			LogText("Session %d (%d)\n",pInfos[i].SessionId,pInfos[i].State);
			if (pInfos[i].State==WTSActive)
			{
				StartStartMenu(pInfos[i].SessionId);
			}
		}
		WTSFreeMemory(pInfos);
	}
	g_hServiceStatus=RegisterServiceCtrlHandlerEx(g_ServiceName,ServiceHandlerEx,NULL);
	if (g_hServiceStatus)
	{
		g_ServiceStatus.dwServiceType=SERVICE_WIN32;
		g_ServiceStatus.dwControlsAccepted=SERVICE_ACCEPT_STOP|SERVICE_ACCEPT_SHUTDOWN|SERVICE_ACCEPT_PAUSE_CONTINUE|SERVICE_ACCEPT_SESSIONCHANGE;
		g_ServiceStatus.dwWin32ExitCode=0;
		g_ServiceStatus.dwServiceSpecificExitCode=0;
		g_ServiceStatus.dwCurrentState=SERVICE_RUNNING;
		g_ServiceStatus.dwCheckPoint=0;
		g_ServiceStatus.dwWaitHint=0;
		SetServiceStatus(g_hServiceStatus, &g_ServiceStatus);
	}
}

#ifndef BUILD_SETUP
// Allow the service to install and uninstall itself during development
static void InstallService( void )
{
	wchar_t path[_MAX_PATH];
	GetModuleFileName(NULL,path,_countof(path));
	SC_HANDLE hManager=OpenSCManager(NULL,NULL,SC_MANAGER_CREATE_SERVICE); 
	if (hManager)
	{
		SC_HANDLE hService=CreateService(hManager,g_ServiceName,g_ServiceName,SERVICE_ALL_ACCESS,SERVICE_WIN32_OWN_PROCESS,SERVICE_AUTO_START,SERVICE_ERROR_NORMAL,path,L"UIGroup",NULL,NULL,NULL,NULL);
		if (hService)
		{
			SERVICE_DESCRIPTION desc={L"Launches the start button after logon"};
			ChangeServiceConfig2(hService,SERVICE_CONFIG_DESCRIPTION,&desc);
			CloseServiceHandle(hService);
		}
		CloseServiceHandle(hManager);
	}	
}

static void UninstallService( void )
{
	SC_HANDLE hManager=OpenSCManager(NULL,NULL,SC_MANAGER_ALL_ACCESS);
	if (hManager)
	{
		SC_HANDLE hService=OpenService(hManager,g_ServiceName,SERVICE_ALL_ACCESS);
		if (hService)
		{
			DeleteService(hService);
			CloseServiceHandle(hService);
		}
		CloseServiceHandle(hManager);
	}
}

#endif

// MiniDumpNormal - minimal information
// MiniDumpWithDataSegs - include global variables
// MiniDumpWithFullMemory - include heap
MINIDUMP_TYPE MiniDumpType=MiniDumpNormal;

static DWORD WINAPI SaveCrashDump( void *pExceptionInfo )
{
	HMODULE dbghelp=NULL;
	{
		wchar_t path[_MAX_PATH];
		GetModuleFileName(NULL,path,_countof(path));
		PathRemoveFileSpec(path);

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
				_swprintf(fname,L"%s\\CS_Crash%d.dmp",path,i);
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

int wmain( int argc, const wchar_t *argv[] )
{
#ifndef BUILD_SETUP
	if(argc==2)
	{
		if(wcscmp(L"-install",argv[1])==0)
			InstallService();
		else if (wcscmp(L"-uninstall",argv[1])==0)
			UninstallService();
		return 0;
	}
#endif
	SERVICE_TABLE_ENTRY DispatchTable[]={
		{(wchar_t*)g_ServiceName, ServiceMain},
		{NULL, NULL}
	};
	HKEY hKey;
	if (RegOpenKeyEx(HKEY_LOCAL_MACHINE,L"SOFTWARE\\IvoSoft\\ClassicShell",0,KEY_READ|KEY_WOW64_64KEY,&hKey)==ERROR_SUCCESS)
	{
		DWORD log;
		DWORD size=sizeof(log);
		if (RegQueryValueEx(hKey,L"LogService",0,NULL,(BYTE*)&log,&size)==ERROR_SUCCESS && log)
		{
			GetModuleFileName(NULL,g_LogName,_countof(g_LogName));
			PathRemoveFileSpec(g_LogName);
			PathAppend(g_LogName,L"service.log");
			LogText("Starting service\n");
		}

		DWORD dump;
		size=sizeof(dump);

		if (RegQueryValueEx(hKey,L"CrashDump",0,NULL,(BYTE*)&dump,&size)==ERROR_SUCCESS && dump>0)
		{
			if (dump==1) MiniDumpType=MiniDumpNormal;
			if (dump==2) MiniDumpType=MiniDumpWithDataSegs;
			if (dump==3) MiniDumpType=MiniDumpWithFullMemory;
			SetUnhandledExceptionFilter(TopLevelFilter);
		}
		RegCloseKey(hKey);
	}

	StartServiceCtrlDispatcher(DispatchTable);
	return 0;
}
