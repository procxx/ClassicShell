#include <windows.h>
#include <WtsApi32.h>
#include <shlwapi.h>

static const wchar_t *g_ServiceName=L"ClassicShellService";
static SERVICE_STATUS_HANDLE g_hServiceStatus;
static SERVICE_STATUS g_ServiceStatus;

static DWORD WINAPI ServiceHandlerEx(DWORD dwControl, DWORD dwEventType, LPVOID lpEventData, LPVOID lpContext)
{
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
				// run the classic start menu on logon
				HANDLE hUser;
				if (WTSQueryUserToken(((WTSSESSION_NOTIFICATION*)lpEventData)->dwSessionId,&hUser))
				{
					STARTUPINFO startupInfo={sizeof(STARTUPINFO),NULL,L"Winsta0\\Default"};
					PROCESS_INFORMATION processInfo;
					wchar_t path[_MAX_PATH];
					GetModuleFileName(NULL,path,_countof(path));
					PathRemoveFileSpec(path);
					PathAppend(path,L"ClassicStartMenu.exe -startup");
					if(CreateProcessAsUser(hUser,NULL,path,NULL,NULL,TRUE,NORMAL_PRIORITY_CLASS,NULL,NULL,&startupInfo,&processInfo))
					{
						CloseHandle(processInfo.hProcess);
						CloseHandle(processInfo.hThread);
					}
					CloseHandle(hUser);
				}
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
	StartServiceCtrlDispatcher(DispatchTable);
	return 0;
}
