// Classic Shell (c) 2009-2012, Ivo Beltchev
// The sources for Classic Shell are distributed under the MIT open source license

#define STRICT_TYPED_ITEMIDS
#include <windows.h>
#include <atlbase.h>
#include <atlwin.h>
#include <atlstr.h>
#include "resource.h"
#include "StringUtils.h"
#include "Settings.h"
#include "SettingsUIHelper.h"
#include "ResourceHelper.h"
#include "Translations.h"
#include <wininet.h>

// Manifest to enable the 6.0 common controls
#pragma comment(linker, \
	"\"/manifestdependency:type='Win32' "\
	"name='Microsoft.Windows.Common-Controls' "\
	"version='6.0.0.0' "\
	"processorArchitecture='*' "\
	"publicKeyToken='6595b64144ccf1df' "\
	"language='*'\"")


void ClosingSettings( HWND hWnd, int flags, int command )
{
}

void UpdateSettings( void )
{
}

const wchar_t *GetDocRelativePath( void )
{
	return NULL;
}

static HINSTANCE g_Instance;

static int g_LoadDialogs[]=
{
	IDD_UPDATE,
	0
};

static CSetting g_Settings[]={
{L"Update",CSetting::TYPE_GROUP},
	{L"Language",CSetting::TYPE_STRING,0,0,L"",CSetting::FLAG_SHARED},
	{L"Update",CSetting::TYPE_BOOL,0,0,1,CSetting::FLAG_SHARED},
	{L"RemindedVersion",CSetting::TYPE_INT,0,0,0,CSetting::FLAG_SHARED},

	{NULL}
};

const int SETTING_UPDATE=2;
const int SETTING_REMINDED=3;

///////////////////////////////////////////////////////////////////////////////

class CUpdateDlg: public CResizeableDlg<CUpdateDlg>
{
public:
	BEGIN_MSG_MAP( CUpdateDlg )
		MESSAGE_HANDLER( WM_INITDIALOG, OnInitDialog )
		MESSAGE_HANDLER( WM_DESTROY, OnDestroy )
		MESSAGE_HANDLER( WM_SIZE, OnSize )
		MESSAGE_HANDLER( WM_GETMINMAXINFO, OnGetMinMaxInfo )
		MESSAGE_HANDLER( WM_CTLCOLORSTATIC, OnColorStatic )
		COMMAND_HANDLER( IDC_CHECKAUTOCHECK, BN_CLICKED, OnCheckAuto )
		COMMAND_HANDLER( IDC_BUTTONCHECKNOW, BN_CLICKED, OnCheckNow )
		COMMAND_HANDLER( IDC_BUTTONDOWNLOAD, BN_CLICKED, OnDownload )
		COMMAND_HANDLER( IDC_CHECKDONT, BN_CLICKED, OnDontRemind )
		COMMAND_HANDLER( IDOK, BN_CLICKED, OnOK )
		COMMAND_HANDLER( IDCANCEL, BN_CLICKED, OnCancel )
	END_MSG_MAP()

	BEGIN_RESIZE_MAP
		RESIZE_CONTROL(IDC_STATICLATEST,MOVE_SIZE_X)
		RESIZE_CONTROL(IDC_EDITTEXT,MOVE_SIZE_X|MOVE_SIZE_Y)
		RESIZE_CONTROL(IDC_BUTTONDOWNLOAD,MOVE_MOVE_Y)
		RESIZE_CONTROL(IDC_CHECKDONT,MOVE_MOVE_Y)
		RESIZE_CONTROL(IDOK,MOVE_MOVE_X|MOVE_MOVE_Y)
		RESIZE_CONTROL(IDCANCEL,MOVE_MOVE_X|MOVE_MOVE_Y)
	END_RESIZE_MAP

	CUpdateDlg( void ) { m_NewVersion=0xFFFFFFFF; }
	void Run( void );
	void UpdateData( void );

protected:
	// Handler prototypes:
	//  LRESULT MessageHandler(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
	//  LRESULT CommandHandler(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled);
	//  LRESULT NotifyHandler(int idCtrl, LPNMHDR pnmh, BOOL& bHandled);
	LRESULT OnInitDialog( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled );
	LRESULT OnDestroy( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled );
	LRESULT OnSize( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled );
	LRESULT OnOK( WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled );
	LRESULT OnCancel( WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled );
	LRESULT OnColorStatic( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled );
	LRESULT OnCheckAuto( WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled );
	LRESULT OnCheckNow( WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled );
	LRESULT OnDownload( WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled );
	LRESULT OnDontRemind( WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled );

private:
	CString m_DownloadUrl;
	CString m_News;
	CWindow m_Tooltip;
	DWORD m_Version;
	DWORD m_NewVersion;
	DWORD m_RemindedVersion;
	HFONT m_Font;

	void UpdateUI( void );

	static void NewVersionCallback( DWORD newVersion, CString downloadUrl, CString news );
};

static CUpdateDlg g_UpdateDlg;

LRESULT CUpdateDlg::OnInitDialog( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled )
{
	CResizeableDlg<CUpdateDlg>::InitResize(MOVE_MODAL);
	m_Version=GetVersionEx(_AtlBaseModule.GetModuleInstance());

	HICON icon=(HICON)LoadImage(g_Instance,MAKEINTRESOURCE(IDI_CLASSICSHELL),IMAGE_ICON,GetSystemMetrics(SM_CXICON),GetSystemMetrics(SM_CYICON),LR_DEFAULTCOLOR);
	SendMessage(WM_SETICON,ICON_BIG,(LPARAM)icon);
	icon=(HICON)LoadImage(g_Instance,MAKEINTRESOURCE(IDI_CLASSICSHELL),IMAGE_ICON,GetSystemMetrics(SM_CXSMICON),GetSystemMetrics(SM_CYSMICON),LR_DEFAULTCOLOR);
	SendMessage(WM_SETICON,ICON_SMALL,(LPARAM)icon);

	HDC hdc=::GetDC(NULL);
	int dpi=GetDeviceCaps(hdc,LOGPIXELSY);
	::ReleaseDC(NULL,hdc);
	m_Font=CreateFont(-9*dpi/72,0,0,0,FW_NORMAL,0,0,0,DEFAULT_CHARSET,OUT_DEFAULT_PRECIS,CLIP_DEFAULT_PRECIS,DEFAULT_QUALITY,FIXED_PITCH,L"Consolas");
	if (m_Font)
		GetDlgItem(IDC_EDITTEXT).SetFont(m_Font);
	m_Tooltip.Create(TOOLTIPS_CLASS,m_hWnd,NULL,NULL,WS_POPUP|TTS_NOPREFIX);

	bool check=true;
	if (g_Settings[SETTING_UPDATE].value.vt==VT_I4)
		check=g_Settings[SETTING_UPDATE].value.intVal!=0;
	CheckDlgButton(IDC_CHECKAUTOCHECK,check?BST_CHECKED:BST_UNCHECKED);
	GetDlgItem(IDC_CHECKAUTOCHECK).EnableWindow(!(g_Settings[SETTING_UPDATE].flags&CSetting::FLAG_LOCKED_MASK));
	GetDlgItem(IDC_BUTTONCHECKNOW).EnableWindow(!(g_Settings[SETTING_UPDATE].flags&CSetting::FLAG_LOCKED_MASK) || check);
	m_RemindedVersion=m_Version;
	if (g_Settings[SETTING_REMINDED].value.vt==VT_I4)
		m_RemindedVersion=g_Settings[SETTING_REMINDED].value.intVal;
	UpdateUI();

	return TRUE;
}

LRESULT CUpdateDlg::OnDestroy( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled )
{
	if (m_Font) DeleteObject(m_Font);
	return 0;
}

LRESULT CUpdateDlg::OnSize( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled )
{
	CResizeableDlg<CUpdateDlg>::OnSize();
	return 0;
}

LRESULT CUpdateDlg::OnOK( WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled )
{
	SaveSettings();
	DestroyWindow();
	return 0;
}

LRESULT CUpdateDlg::OnCancel( WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled )
{
	DestroyWindow();
	return 0;
}

LRESULT CUpdateDlg::OnColorStatic( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled )
{
	if ((m_NewVersion==0 || m_Version<m_NewVersion) && lParam==(LPARAM)GetDlgItem(IDC_STATICLATEST).m_hWnd)
	{
		HDC hdc=(HDC)wParam;
		SetTextColor(hdc,0xFF);
		SetBkMode(hdc,TRANSPARENT);
		return (LRESULT)GetSysColorBrush(COLOR_3DFACE);
	}
	bHandled=FALSE;
	return 0;
}

LRESULT CUpdateDlg::OnCheckAuto( WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled )
{
	CSettingsLockWrite lock;
	bool check=IsDlgButtonChecked(IDC_CHECKAUTOCHECK)==BST_CHECKED;
	g_Settings[SETTING_UPDATE].value=CComVariant(check?1:0);
	g_Settings[SETTING_UPDATE].flags&=~CSetting::FLAG_DEFAULT;
	UpdateUI();
	return 0;
}

LRESULT CUpdateDlg::OnCheckNow( WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled )
{
	UpdateData();
	UpdateUI();
	return 0;
}

LRESULT CUpdateDlg::OnDownload( WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled )
{
	if (!m_DownloadUrl.IsEmpty())
	{
		ShellExecute(GetWindow(GA_ROOT),NULL,m_DownloadUrl,NULL,NULL,SW_SHOWNORMAL);
	}
	return 0;
}

LRESULT CUpdateDlg::OnDontRemind( WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled )
{
	if (IsDlgButtonChecked(IDC_CHECKDONT)==BST_CHECKED)
	{
		m_RemindedVersion=m_NewVersion;
	}
	else
	{
		m_RemindedVersion=m_NewVersion-1;
	}

	CSettingsLockWrite lock;
	g_Settings[SETTING_REMINDED].value=CComVariant((int)m_RemindedVersion);
	g_Settings[SETTING_REMINDED].flags&=~CSetting::FLAG_DEFAULT;
	return 0;
}

void CUpdateDlg::NewVersionCallback( DWORD newVersion, CString downloadUrl, CString news )
{
	g_UpdateDlg.m_NewVersion=newVersion;
	g_UpdateDlg.m_DownloadUrl=downloadUrl;
	g_UpdateDlg.m_News=news;
}

void CUpdateDlg::UpdateData( void )
{
	if (!CheckForNewVersion(CHECK_UPDATE,NewVersionCallback))
	{
		m_NewVersion=0;
		m_DownloadUrl.Empty();
		m_News.Empty();
	}
}

void CUpdateDlg::UpdateUI( void )
{
	if (m_NewVersion!=0 && m_NewVersion!=0xFFFFFFFF)
	{
		if (m_Version>=m_NewVersion)
		{
			SetDlgItemText(IDC_STATICLATEST,LoadStringEx(IDS_UPDATED));
			SetDlgItemText(IDC_EDITTEXT,L"");
			GetDlgItem(IDC_EDITTEXT).ShowWindow(SW_HIDE);
			GetDlgItem(IDC_BUTTONDOWNLOAD).ShowWindow(SW_HIDE);
			GetDlgItem(IDC_CHECKDONT).ShowWindow(SW_HIDE);
		}
		else
		{
			SetDlgItemText(IDC_STATICLATEST,LoadStringEx(IDS_OUTOFDATE));
			SetDlgItemText(IDC_EDITTEXT,m_News);
			GetDlgItem(IDC_EDITTEXT).ShowWindow(SW_SHOW);
			GetDlgItem(IDC_BUTTONDOWNLOAD).ShowWindow(SW_SHOW);
			bool check=true;
			if (g_Settings[SETTING_UPDATE].value.vt==VT_I4)
				check=g_Settings[SETTING_UPDATE].value.intVal!=0;
			GetDlgItem(IDC_CHECKDONT).ShowWindow(check?SW_SHOW:SW_HIDE);
			CheckDlgButton(IDC_CHECKDONT,m_RemindedVersion>=m_NewVersion?BST_CHECKED:BST_UNCHECKED);
			TOOLINFO tool={sizeof(tool),TTF_SUBCLASS|TTF_IDISHWND,m_hWnd,(UINT_PTR)GetDlgItem(IDC_BUTTONDOWNLOAD).m_hWnd};
			tool.lpszText=(LPWSTR)(LPCWSTR)m_DownloadUrl;
			m_Tooltip.SendMessage(TTM_ADDTOOL,0,(LPARAM)&tool);
		}
	}
	else
	{
		SetDlgItemText(IDC_STATICLATEST,(m_NewVersion==0)?LoadStringEx(IDS_UPDATE_FAIL):L"");
		SetDlgItemText(IDC_EDITTEXT,L"");
		GetDlgItem(IDC_EDITTEXT).ShowWindow(SW_HIDE);
		GetDlgItem(IDC_BUTTONDOWNLOAD).ShowWindow(SW_HIDE);
		GetDlgItem(IDC_CHECKDONT).ShowWindow(SW_HIDE);
	}
	Invalidate();
}

void CUpdateDlg::Run( void )
{
	DLGTEMPLATE *pTemplate=LoadDialogEx(IDD_UPDATE);
	Create(NULL,pTemplate);
	MSG msg;
	while (m_hWnd && GetMessage(&msg,0,0,0))
	{
		if ((msg.hwnd==m_hWnd || IsChild(msg.hwnd)) && IsDialogMessage(&msg))
			continue;
		TranslateMessage(&msg);
		DispatchMessage(&msg);
	}
}

///////////////////////////////////////////////////////////////////////////////

LRESULT CALLBACK SubclassBalloonProc( HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData )
{
	if (uMsg==WM_LBUTTONUP)
	{
		LRESULT res=DefSubclassProc(hWnd,uMsg,wParam,lParam);
		if (SendMessage(hWnd,TTM_GETCURRENTTOOL,0,0))
		{
			TOOLINFO tool={sizeof(tool)};
			tool.uId=1;
			SendMessage(hWnd,TTM_TRACKACTIVATE,FALSE,(LPARAM)&tool);
			g_UpdateDlg.Run();
		}
		return res;
	}
	if (uMsg==WM_MOUSEACTIVATE)
		return MA_NOACTIVATE;
	return DefSubclassProc(hWnd,uMsg,wParam,lParam);
}

///////////////////////////////////////////////////////////////////////////////

int WINAPI wWinMain( HINSTANCE hInstance, HINSTANCE hPrevInstance, LPTSTR lpstrCmdLine, int nCmdShow )
{
	INITCOMMONCONTROLSEX init={sizeof(init),ICC_STANDARD_CLASSES};
	InitCommonControlsEx(&init);
	SetProcessDPIAware();

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
	Sprintf(mutexName,_countof(mutexName),L"ClassicShellUpdate.Mutex.%s.%s",userName,deskName);
	free(deskName);

	HANDLE hMutex=CreateMutex(NULL,TRUE,mutexName);
	if (GetLastError()==ERROR_ALREADY_EXISTS || GetLastError()==ERROR_ACCESS_DENIED)
		return 0;

	g_Instance=hInstance;
	InitSettings(g_Settings,COMPONENT_UPDATE);
	CString language=GetSettingString(L"Language");
	ParseTranslations(NULL,language);

	g_Instance=hInstance;
	wchar_t path[_MAX_PATH];
	GetModuleFileName(hInstance,path,_countof(path));
	*PathFindFileName(path)=0;
	HINSTANCE resInstance=NULL;
	if (!language.IsEmpty())
	{
		wchar_t fname[_MAX_PATH];
		Sprintf(fname,_countof(fname),L"%s%s.dll",path,language);
		resInstance=LoadLibraryEx(fname,NULL,LOAD_LIBRARY_AS_DATAFILE|LOAD_LIBRARY_AS_IMAGE_RESOURCE);
	}
	else
	{
		wchar_t languages[100]={0};
		ULONG size=4; // up to 4 languages
		ULONG len=_countof(languages);
		GetThreadPreferredUILanguages(MUI_LANGUAGE_NAME,&size,languages,&len);

		for (const wchar_t *language=languages;*language;language+=Strlen(language)+1)
		{
			wchar_t fname[_MAX_PATH];
			Sprintf(fname,_countof(fname),L"%s%s.dll",path,language);
			resInstance=LoadLibraryEx(fname,NULL,LOAD_LIBRARY_AS_DATAFILE|LOAD_LIBRARY_AS_IMAGE_RESOURCE);
			if (resInstance)
				break;
		}
	}

	if (resInstance && GetVersionEx(resInstance)!=GetVersionEx(g_Instance))
	{
		FreeLibrary(resInstance);
		resInstance=NULL;
	}
	LoadTranslationResources(g_Instance,resInstance,g_LoadDialogs);

	if (resInstance)
		FreeLibrary(resInstance);

	int time0=timeGetTime();

	if (wcsstr(lpstrCmdLine,L"-popup")!=NULL)
	{
		g_UpdateDlg.UpdateData();
		int sleep=timeGetTime()-time0;
		if (sleep>0)
			Sleep(sleep);
		HWND balloon=CreateWindowEx(WS_EX_TOPMOST|WS_EX_TOOLWINDOW|(IsLanguageRTL()?WS_EX_LAYOUTRTL:0),TOOLTIPS_CLASS,NULL,WS_POPUP|TTS_CLOSE|TTS_NOPREFIX,0,0,0,0,NULL,NULL,g_Instance,NULL);
		SendMessage(balloon,TTM_SETMAXTIPWIDTH,0,500);
		TOOLINFO tool={sizeof(tool),TTF_ABSOLUTE|TTF_TRANSPARENT|TTF_TRACK|(IsLanguageRTL()?TTF_RTLREADING:0)};
		tool.uId=1;
		CString message=LoadStringEx(IDS_NEWVERSION);
		tool.lpszText=(wchar_t*)(const wchar_t*)message;
		SendMessage(balloon,TTM_ADDTOOL,0,(LPARAM)&tool);
		SendMessage(balloon,TTM_SETTITLE,(WPARAM)LoadIcon(g_Instance,MAKEINTRESOURCE(IDI_CLASSICSHELL)),(LPARAM)(const wchar_t*)LoadStringEx(IDS_UPDATE_TITLE));
		APPBARDATA appbar={sizeof(appbar)};
		SHAppBarMessage(ABM_GETTASKBARPOS,&appbar);
		MONITORINFO info={sizeof(info)};
		GetMonitorInfo(MonitorFromWindow(appbar.hWnd,MONITOR_DEFAULTTOPRIMARY),&info);
		SendMessage(balloon,TTM_TRACKPOSITION,0,0);
		SendMessage(balloon,TTM_TRACKACTIVATE,TRUE,(LPARAM)&tool);
		RECT rc;
		GetWindowRect(balloon,&rc);
		LONG pos;
		if (appbar.uEdge==ABE_LEFT)
			pos=MAKELONG(info.rcWork.left,info.rcWork.bottom-rc.bottom+rc.top);
		else if (appbar.uEdge==ABE_RIGHT)
			pos=MAKELONG(info.rcWork.right-rc.right+rc.left,info.rcWork.bottom-rc.bottom+rc.top);
		else if (appbar.uEdge==ABE_TOP)
			pos=MAKELONG(IsLanguageRTL()?info.rcWork.left:info.rcWork.right-rc.right+rc.left,info.rcWork.top);
		else
			pos=MAKELONG(IsLanguageRTL()?info.rcWork.left:info.rcWork.right-rc.right+rc.left,info.rcWork.bottom-rc.bottom+rc.top);
		SendMessage(balloon,TTM_TRACKPOSITION,0,pos);
		SetWindowSubclass(balloon,SubclassBalloonProc,0,'CLSH');
		PlaySound(L"SystemNotification",NULL,SND_APPLICATION|SND_ALIAS|SND_ASYNC|SND_NODEFAULT|SND_SYSTEM);
		int time0=timeGetTime();
		while (IsWindowVisible(balloon))
		{
			if (time0 && (timeGetTime()-time0)>=15000)
			{
				time0=0;
				TOOLINFO tool={sizeof(tool)};
				tool.uId=1;
				SendMessage(balloon,TTM_TRACKACTIVATE,FALSE,(LPARAM)&tool);
			}
			MSG msg;
			while (PeekMessage(&msg,0,0,0,PM_REMOVE))
			{
				TranslateMessage(&msg);
				DispatchMessage(&msg);
			}
			Sleep(10);
		}
	}
	else
	{
		g_UpdateDlg.Run();
	}
	return 0;
}
