// Classic Shell (c) 2009-2012, Ivo Beltchev
// The sources for Classic Shell are distributed under the MIT open source license

#include "stdafx.h"
#include "resource.h"
#include "Translations.h"
#include "Settings.h"
#include "ResourceHelper.h"
#include "IconManager.h"
#include "ClassicStartMenuDLL.h"
#include "ClassicStartButton.h"
#include "dllmain.h"
#include <uxtheme.h>
#include <vsstyle.h>
#include <dwmapi.h>

static int START_ICON_SIZE=16;
const int START_BUTTON_PADDING=3;
const int START_BUTTON_OFFSET=2;
const int START_TEXT_PADDING=2;

// CStartButton - implementation of a start button (for Windows 8)
class CStartButton: public CWindowImpl<CStartButton>
{
public:
	DECLARE_WND_CLASS_EX(L"ClassicShell.CStartButton",CS_DBLCLKS,COLOR_MENU)
	CStartButton( void );

	// message handlers
	BEGIN_MSG_MAP( CStartButton )
		MESSAGE_HANDLER( WM_CREATE, OnCreate )
		MESSAGE_HANDLER( WM_DESTROY, OnDestroy )
		MESSAGE_HANDLER( WM_CLOSE, OnClose )
		MESSAGE_HANDLER( WM_MOUSEACTIVATE, OnMouseActivate )
		MESSAGE_HANDLER( WM_MOUSEMOVE, OnMouseMove )
		MESSAGE_HANDLER( WM_ERASEBKGND, OnEraseBkgnd )
		MESSAGE_HANDLER( WM_TIMER, OnTimer )
		MESSAGE_HANDLER( WM_SETTINGCHANGE, OnSettingChange )
		MESSAGE_HANDLER( WM_THEMECHANGED, OnThemeChanged )
	END_MSG_MAP()

	void SetPressed( bool bPressed );
	void UpdateButton( void );
	void TaskBarMouseMove( void );

	SIZE GetSize( void ) const { return m_Size; }
	bool GetSmallIcons( void ) const { return m_bSmallIcons; }

protected:
	LRESULT OnCreate( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled );
	LRESULT OnDestroy( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled );
	LRESULT OnClose( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled ) { return 0; }
	LRESULT OnEraseBkgnd( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled ) { return 1; }
	LRESULT OnMouseActivate( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled ) { return MA_NOACTIVATE; }
	LRESULT OnMouseMove( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled );
	LRESULT OnTimer( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled );
	LRESULT OnSettingChange( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled );
	LRESULT OnThemeChanged( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled );

private:
	enum { TIMER_BLEND=1, TIMER_LEAVE=2 };

	SIZE m_Size;
	HBITMAP m_Bitmap, m_Blendmap;
	unsigned int *m_Bits, *m_BlendBits;
	HICON m_Icon;
	HFONT m_Font;
	bool m_bHot, m_bPressed;
	bool m_bTrackMouse;
	bool m_bClassic;
	bool m_bRTL;
	bool m_bFlipped; // the image is flipped
	bool m_bSmallIcons;
	int m_HotBlend; // 0..100
	CWindow m_Tooltip;
	HTHEME m_Theme;

	void LoadBitmap( void );
	void SetHot( bool bHot );
};

CStartButton::CStartButton( void )
{
	m_Bitmap=m_Blendmap=NULL;
	m_Icon=NULL;
	m_Font=NULL;
	m_bHot=m_bPressed=false;
	m_bTrackMouse=false;
	m_bClassic=false;
	m_bRTL=false;
}

LRESULT CStartButton::OnCreate( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled )
{
	m_bRTL=((CREATESTRUCT*)lParam)->lpCreateParams!=NULL;
	m_bSmallIcons=IsTaskbarSmallIcons();
	std::vector<HMODULE> modules;
	m_Icon=NULL;
	START_ICON_SIZE=0;
	CString iconPath=GetSettingString(L"StartButtonIcon");
	if (_wcsicmp(iconPath,L"none")!=0)
	{
		START_ICON_SIZE=GetSystemMetrics(m_bSmallIcons?SM_CXSMICON:SM_CXICON);
		m_Icon=LoadIcon(START_ICON_SIZE,iconPath,modules);
		for (std::vector<HMODULE>::const_iterator it=modules.begin();it!=modules.end();++it)
			FreeLibrary(*it);
		if (!m_Icon)
			m_Icon=(HICON)LoadImage(g_Instance,MAKEINTRESOURCE(IDI_APPICON),IMAGE_ICON,START_ICON_SIZE,START_ICON_SIZE,LR_DEFAULTCOLOR);
	}
	int dpi=CIconManager::GetDPI();
	m_Font=CreateFont(10*dpi/72,0,0,0,FW_BOLD,0,0,0,DEFAULT_CHARSET,OUT_DEFAULT_PRECIS,CLIP_DEFAULT_PRECIS,DEFAULT_QUALITY,DEFAULT_PITCH,L"Tahoma");
	int val=1;
	DwmSetWindowAttribute(m_hWnd,DWMWA_EXCLUDED_FROM_PEEK,&val,sizeof(val));
	val=DWMFLIP3D_EXCLUDEABOVE;
	DwmSetWindowAttribute(m_hWnd,DWMWA_FLIP3D_POLICY,&val,sizeof(val));
	LoadBitmap();
	m_Tooltip=CreateWindowEx(WS_EX_TOPMOST|WS_EX_TOOLWINDOW|WS_EX_TRANSPARENT|(m_bRTL?WS_EX_LAYOUTRTL:0),TOOLTIPS_CLASS,NULL,WS_POPUP|TTS_NOPREFIX|TTS_ALWAYSTIP,0,0,0,0,NULL,NULL,g_Instance,NULL);
	OnThemeChanged(WM_THEMECHANGED,0,0,bHandled);
	m_bPressed=true;
	SetPressed(false);
	bHandled=FALSE;
	return 0;
}

LRESULT CStartButton::OnDestroy( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled )
{
	if (m_Bitmap) DeleteObject(m_Bitmap);
	if (m_Blendmap) DeleteObject(m_Blendmap);
	if (m_Icon) DestroyIcon(m_Icon);
	if (m_Font) DeleteObject(m_Font);
	if (m_Theme) CloseThemeData(m_Theme);
	m_Tooltip.DestroyWindow();
	KillTimer(TIMER_BLEND);
	bHandled=FALSE;
	return 0;
}

void CStartButton::UpdateButton( void )
{
	BLENDFUNCTION func={AC_SRC_OVER,0,255,AC_SRC_ALPHA};

	HDC hSrc=CreateCompatibleDC(NULL);
	RECT rc;
	GetWindowRect(&rc);
	SIZE size={rc.right-rc.left,rc.bottom-rc.top};
	if (m_bClassic)
	{
		if (m_bRTL)
			SetLayout(hSrc,LAYOUT_RTL);
		HBITMAP bmp0=(HBITMAP)SelectObject(hSrc,m_Blendmap);
		RECT rc={0,0,m_Size.cx,m_Size.cy};
		FillRect(hSrc,&rc,(HBRUSH)GetStockObject(BLACK_BRUSH));
		InflateRect(&rc,-START_BUTTON_OFFSET,-START_BUTTON_OFFSET);
		int offset=0;
		if (m_Theme)
		{
			int state=m_bPressed?PBS_PRESSED:(m_bHot?PBS_HOT:PBS_NORMAL);
			DrawThemeBackground(m_Theme,hSrc,BP_PUSHBUTTON,state,&rc,NULL);
		}
		else
		{
			DrawFrameControl(hSrc,&rc,DFC_BUTTON,DFCS_BUTTONPUSH|(m_bPressed?DFCS_PUSHED:0));
			offset=m_bPressed?1:0;
		}
		if (m_Icon)
			DrawIconEx(hSrc,START_BUTTON_PADDING+START_BUTTON_OFFSET+offset,(m_Size.cy-START_ICON_SIZE)/2+offset,m_Icon,0,0,0,NULL,DI_NORMAL|DI_NOMIRROR);
		rc.left+=START_BUTTON_PADDING+START_ICON_SIZE+START_TEXT_PADDING+offset;
		rc.top+=START_BUTTON_PADDING+offset;
		rc.right-=START_BUTTON_PADDING+START_TEXT_PADDING-offset;
		rc.bottom-=START_BUTTON_PADDING-offset;
		HFONT font0=(HFONT)SelectObject(hSrc,m_Font);
		SetTextColor(hSrc,GetSysColor(COLOR_BTNTEXT));
		SetBkMode(hSrc,TRANSPARENT);
		CString startStr=GetSettingString(L"StartButtonText");
		const wchar_t *startText=startStr;
		if (startText[0]=='$')
			startText=FindTranslation(startText+1,L"Start");
		DrawText(hSrc,startText,-1,&rc,DT_NOPREFIX|DT_SINGLELINE|DT_VCENTER);
		SelectObject(hSrc,bmp0);
		// mark the button pixels as opaque
		for (int y=START_BUTTON_OFFSET;y<m_Size.cy-START_BUTTON_OFFSET;y++)
			for (int x=START_BUTTON_OFFSET;x<m_Size.cx-START_BUTTON_OFFSET;x++)
				m_BlendBits[y*m_Size.cx+x]|=0xFF000000;
		SelectObject(hSrc,m_Blendmap);

		POINT pos={0,0};
		UpdateLayeredWindow(m_hWnd,NULL,NULL,&size,hSrc,&pos,0,&func,ULW_ALPHA);
		SelectObject(hSrc,font0);
		SelectObject(hSrc,bmp0);
	}
	else
	{
		int image=-1;
		if (m_bPressed) image=2;
		else if (m_HotBlend==0) image=0;
		else if (m_HotBlend==100) image=1;
		if (image!=-1)
		{
			HBITMAP bmp0=(HBITMAP)SelectObject(hSrc,m_Bitmap);
			POINT pos={0,image*m_Size.cy};
			UpdateLayeredWindow(m_hWnd,NULL,NULL,&size,hSrc,&pos,0,&func,ULW_ALPHA);
			SelectObject(hSrc,bmp0);
		}
		else if (m_Bits)
		{
			// blend the two images
			int n=m_Size.cx*m_Size.cy;
			int n2=m_bFlipped?n*2:0;
			for (int i=0;i<n;i++)
			{
				unsigned int pixel1=m_Bits[i+n2];
				unsigned int pixel2=m_Bits[i+n];
				int a1=(pixel1>>24);
				int r1=(pixel1>>16)&255;
				int g1=(pixel1>>8)&255;
				int b1=(pixel1)&255;
				int a2=(pixel2>>24);
				int r2=(pixel2>>16)&255;
				int g2=(pixel2>>8)&255;
				int b2=(pixel2)&255;
				int a=a1+(a2-a1)*m_HotBlend/100;
				int r=r1+(r2-r1)*m_HotBlend/100;
				int g=g1+(g2-g1)*m_HotBlend/100;
				int b=b1+(b2-b1)*m_HotBlend/100;
				m_BlendBits[i]=(a<<24)|(r<<16)|(g<<8)|b;
			}
			HBITMAP bmp0=(HBITMAP)SelectObject(hSrc,m_Blendmap);
			POINT pos={0,0};
			UpdateLayeredWindow(m_hWnd,NULL,NULL,&size,hSrc,&pos,0,&func,ULW_ALPHA);
			SelectObject(hSrc,bmp0);
		}
	}
	DeleteDC(hSrc);
}

void CStartButton::SetHot( bool bHot )
{
	if (m_bHot!=bHot)
	{
		m_bHot=bHot;
		if (!m_bPressed)
		{
			SetTimer(TIMER_BLEND,30);
		}
	}
}

LRESULT CStartButton::OnMouseMove( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled )
{
	SetTimer(TIMER_LEAVE,30);
	SetHot(true);
	return 0;
}

void CStartButton::TaskBarMouseMove( void )
{
	SetHot(true);
	SetTimer(CStartButton::TIMER_LEAVE,30);
}

LRESULT CStartButton::OnTimer( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled )
{
	if (wParam==TIMER_BLEND)
	{
		int blend=m_HotBlend+(m_bHot?10:-10);
		if (blend<0) blend=0;
		if (blend>100) blend=100;
		if (blend!=m_HotBlend)
		{
			m_HotBlend=blend;
			UpdateButton();
		}
		else
			KillTimer(TIMER_BLEND);
	}
	else if (wParam==TIMER_LEAVE)
	{
		CPoint pt(GetMessagePos());
		if (WindowFromPoint(pt)!=m_hWnd && !PointAroundStartButton())
		{
			KillTimer(TIMER_LEAVE);
			SetHot(false);
		}
	}
	else
		bHandled=FALSE;
	return 0;
}

LRESULT CStartButton::OnSettingChange( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled )
{
	UpdateButton();
	bHandled=FALSE;
	return 0;
}

LRESULT CStartButton::OnThemeChanged( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled )
{
	if (m_Theme) CloseThemeData(m_Theme);
	m_Theme=NULL;
	HIGHCONTRAST contrast={sizeof(contrast)};
	if (GetWinVersion()>=WIN_VER_WIN8 && SystemParametersInfo(SPI_GETHIGHCONTRAST,sizeof(contrast),&contrast,0) && (contrast.dwFlags&HCF_HIGHCONTRASTON))
	{
		// only use themes on Win8 with high contrast
		m_Theme=OpenThemeData(m_hWnd,L"button");
		UpdateButton();
	}
	return 0;
}

void CStartButton::SetPressed( bool bPressed )
{
	if (m_bPressed!=bPressed)
	{
		m_bPressed=bPressed;
		m_HotBlend=m_bHot?100:0;
		KillTimer(TIMER_BLEND);
		TOOLINFO tool={sizeof(tool),TTF_CENTERTIP|TTF_SUBCLASS|TTF_IDISHWND|TTF_TRANSPARENT|(m_bRTL?TTF_RTLREADING:0),m_hWnd};
		tool.uId=(UINT_PTR)m_hWnd;
		CString startStr=GetSettingString(L"StartButtonTip");
		const wchar_t *startText=startStr;
		if (startText[0]=='$')
			startText=FindTranslation(startText+1,L"Start");
		wchar_t buf[256];
		Strcpy(buf,_countof(buf),startText);
		DoEnvironmentSubst(buf,_countof(buf));
		tool.lpszText=buf;
		m_Tooltip.SendMessage(bPressed?TTM_DELTOOL:TTM_ADDTOOL,0,(LPARAM)&tool);
		UpdateButton();
	}
}

void CStartButton::LoadBitmap( void )
{
	m_Size.cx=m_Size.cy=0;
	if (m_Bitmap) DeleteObject(m_Bitmap);
	if (m_Blendmap) DeleteObject(m_Blendmap);
	m_Bitmap=m_Blendmap=NULL;
	m_Bits=m_BlendBits=NULL;
	bool bDef;
	TStartButtonType buttonType=(TStartButtonType)GetSettingInt(L"StartButtonType",bDef);
	if (bDef)
	{
		bool bClassic;
		if (GetWinVersion()<WIN_VER_WIN8)
			bClassic=!IsAppThemed();
		else
		{
			HIGHCONTRAST contrast={sizeof(contrast)};
			bClassic=(SystemParametersInfo(SPI_GETHIGHCONTRAST,sizeof(contrast),&contrast,0) && (contrast.dwFlags&HCF_HIGHCONTRASTON));
		}
		buttonType=bClassic?START_BUTTON_CLASSIC:START_BUTTON_AERO;
	}
	m_bClassic=(buttonType==START_BUTTON_CLASSIC);
	bool bMetro=(buttonType==START_BUTTON_METRO);
	wchar_t path[_MAX_PATH];
	SIZE size={0,0};
	if (buttonType==START_BUTTON_CUSTOM)
	{
		Strcpy(path,_countof(path),GetSettingString(L"StartButtonPath"));
		DoEnvironmentSubst(path,_countof(path));
		size.cx=GetSettingInt(L"StartButtonSize");
	}
	m_bFlipped=false;
	if (m_bClassic)
	{
		// classic theme
		HDC hdc=CreateCompatibleDC(NULL);
		HFONT font0=(HFONT)SelectObject(hdc,m_Font);
		RECT rc={0,0,0,0};
		CString startStr=GetSettingString(L"StartButtonText");
		const wchar_t *startText=startStr;
		if (startText[0]=='$')
			startText=FindTranslation(startText+1,L"Start");
		DrawText(hdc,startText,-1,&rc,DT_NOPREFIX|DT_SINGLELINE|DT_CALCRECT);
		m_Size.cx=rc.right+START_ICON_SIZE+2*START_TEXT_PADDING+2*START_BUTTON_PADDING+2*START_BUTTON_OFFSET;
		m_Size.cy=rc.bottom;
		if (m_Size.cy<START_ICON_SIZE) m_Size.cy=START_ICON_SIZE;
		m_Size.cy+=2*START_BUTTON_PADDING+2*START_BUTTON_OFFSET;
	}
	else
	{
		bool bResource=false;
		if (*path)
		{
			m_Bitmap=LoadImageFile(path,&size,true);
		}
		if (!m_Bitmap)
		{
			int id;
			int dpi=CIconManager::GetDPI();
			if (dpi<120)
				id=bMetro?IDB_BUTTON96M:IDB_BUTTON96;
			else if (dpi<144)
				id=bMetro?IDB_BUTTON120M:IDB_BUTTON120;
			else if (dpi<180)
				id=bMetro?IDB_BUTTON144M:IDB_BUTTON144;
			else
				id=bMetro?IDB_BUTTON180M:IDB_BUTTON180;
			m_Bitmap=(HBITMAP)LoadImage(g_Instance,MAKEINTRESOURCE(id),IMAGE_BITMAP,0,0,LR_CREATEDIBSECTION);
			bResource=true;
		}
		BITMAP info;
		GetObject(m_Bitmap,sizeof(info),&info);
		m_Size.cx=info.bmWidth;
		m_Size.cy=info.bmHeight/3;
		m_Bits=(unsigned int*)info.bmBits;
		if (bResource)
		{
			// resources are flipped and need to be premultiplied
			m_bFlipped=true;
			int n=info.bmWidth*info.bmHeight;
			for (int i=0;i<n;i++)
			{
				unsigned int &pixel=m_Bits[i];
				int a=(pixel>>24);
				int r=(pixel>>16)&255;
				int g=(pixel>>8)&255;
				int b=(pixel)&255;
				r=(r*a)/255;
				g=(g*a)/255;
				b=(b*a)/255;

				pixel=(a<<24)|(r<<16)|(g<<8)|b;
			}
		}
	}
	if (m_Size.cx>0)
	{
		BITMAPINFO bi={0};
		bi.bmiHeader.biSize=sizeof(BITMAPINFOHEADER);
		bi.bmiHeader.biWidth=m_Size.cx;
		bi.bmiHeader.biHeight=m_bFlipped?m_Size.cy:-m_Size.cy;
		bi.bmiHeader.biPlanes=1;
		bi.bmiHeader.biBitCount=32;
		HDC hdc=CreateCompatibleDC(NULL);
		m_Blendmap=CreateDIBSection(hdc,&bi,DIB_RGB_COLORS,(void**)&m_BlendBits,NULL,0);
		DeleteDC(hdc);
	}
}

static CStartButton g_StartButtonWnd;

HWND CreateStartButton( HWND taskBar, HWND rebar, const RECT &rcTask )
{
	bool bRTL=(GetWindowLong(rebar,GWL_EXSTYLE)&WS_EX_LAYOUTRTL)!=0;
	g_StartButtonWnd.Create(taskBar,NULL,NULL,WS_POPUP,WS_EX_TOPMOST|WS_EX_TOOLWINDOW|WS_EX_LAYERED,0U,(void*)(bRTL?1:0));
	SIZE size=g_StartButtonWnd.GetSize();
	RECT rcButton;
	if (GetWindowLong(rebar,GWL_STYLE)&CCS_VERT)
	{
		rcButton.left=(rcTask.left+rcTask.right-size.cx)/2;
		rcButton.top=rcTask.top;
	}
	else if (bRTL)
	{
		rcButton.left=rcTask.right-size.cx;
		rcButton.top=(rcTask.top+rcTask.bottom-size.cy)/2;
	}
	else
	{
		rcButton.left=rcTask.left;
		rcButton.top=(rcTask.top+rcTask.bottom-size.cy)/2;
	}
	rcButton.right=rcButton.left+size.cx;
	rcButton.bottom=rcButton.top+size.cy;
	g_StartButtonWnd.SetWindowPos(HWND_TOP,&rcButton,SWP_SHOWWINDOW|SWP_NOOWNERZORDER|SWP_NOACTIVATE);

	MONITORINFO info={sizeof(info)};
	GetMonitorInfo(MonitorFromWindow(g_StartButtonWnd,MONITOR_DEFAULTTONEAREST),&info);
	RECT rc;
	IntersectRect(&rc,&rcButton,&info.rcMonitor);
	HRGN rgn=CreateRectRgn(rc.left-rcButton.left,rc.top-rcButton.top,rc.right-rcButton.left,rc.bottom-rcButton.top);
	if (!SetWindowRgn(g_StartButtonWnd,rgn,FALSE))
		DeleteObject(rgn);

	g_StartButtonWnd.UpdateButton();
	return g_StartButtonWnd.m_hWnd;
}

void DestroyStartButton( void )
{
	if (g_StartButtonWnd.m_hWnd)
		g_StartButtonWnd.DestroyWindow();
}

void UpdateStartButton( void )
{
	g_StartButtonWnd.UpdateButton();
}

void PressStartButton( bool bPressed )
{
	if (g_StartButtonWnd.m_hWnd)
		g_StartButtonWnd.SetPressed(bPressed);
}

SIZE GetStartButtonSize( void )
{
	return g_StartButtonWnd.GetSize();
}

bool IsStartButtonSmallIcons( void )
{
	return g_StartButtonWnd.GetSmallIcons();
}

bool IsTaskbarSmallIcons( void )
{
	CRegKey regKey;
	if (regKey.Open(HKEY_CURRENT_USER,L"Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Advanced")!=ERROR_SUCCESS)
		return true;
	DWORD val;
	return regKey.QueryDWORDValue(L"TaskbarSmallIcons",val)!=ERROR_SUCCESS || val;
}

void TaskBarMouseMove( void )
{
	g_StartButtonWnd.TaskBarMouseMove();
}
