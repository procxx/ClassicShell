// Classic Shell (c) 2009-2011, Ivo Beltchev
// The sources for Classic Shell are distributed under the MIT open source license

#include "stdafx.h"
#include "Settings.h"
#include "SettingsUIHelper.h"
#include "ResourceHelper.h"
#include "Translations.h"
#include "resource.h"
#include "dllmain.h"
#include "ClassicIE9DLL.h"
#include <dwmapi.h>
#include <vssym32.h>

///////////////////////////////////////////////////////////////////////////////

static CSetting g_Settings[]={
{L"Basic",CSetting::TYPE_GROUP,IDS_BASIC_SETTINGS},
	{L"EnableSettings",CSetting::TYPE_BOOL,0,0,1,CSetting::FLAG_HIDDEN},
	{L"LogLevel",CSetting::TYPE_INT,0,0,0,CSetting::FLAG_HIDDEN},

{L"TitleBar",CSetting::TYPE_GROUP,IDS_TITLE_SETTINGS},
	{L"ShowCaption",CSetting::TYPE_BOOL,IDS_SHOW_CAPTION,IDS_SHOW_CAPTION_TIP,1,CSetting::FLAG_WARM|CSetting::FLAG_BASIC},
	{L"ShowIcon",CSetting::TYPE_BOOL,IDS_SHOW_ICON,IDS_SHOW_ICON_TIP,1,CSetting::FLAG_WARM|CSetting::FLAG_BASIC,L"ShowCaption"},
	{L"CenterCaption",CSetting::TYPE_BOOL,IDS_CENTER_CAPTION,IDS_CENTER_CAPTION_TIP,0,CSetting::FLAG_WARM|CSetting::FLAG_BASIC,L"ShowCaption"},
	{L"CaptionFont",CSetting::TYPE_FONT,IDS_CAPTION_FONT,IDS_CAPTION_FONT_TIP,L"Segoe UI, normal, 9",CSetting::FLAG_WARM,L"ShowCaption"},
	{L"TextColor",CSetting::TYPE_COLOR,IDS_TEXT_COLOR,IDS_TEXT_COLOR_TIP,0,CSetting::FLAG_WARM,L"ShowCaption"},
	{L"MaxColor",CSetting::TYPE_COLOR,IDS_MAXTEXT_COLOR,IDS_MAXTEXT_COLOR_TIP,0,CSetting::FLAG_WARM|(1<<24),L"ShowCaption"},
	{L"InactiveColor",CSetting::TYPE_COLOR,IDS_INTEXT_COLOR,IDS_INTEXT_COLOR_TIP,0,CSetting::FLAG_WARM|(2<<24),L"ShowCaption"},
	{L"InactiveMaxColor",CSetting::TYPE_COLOR,IDS_MAXINTEXT_COLOR,IDS_MAXINTEXT_COLOR_TIP,0,CSetting::FLAG_WARM|(3<<24),L"ShowCaption"},
	{L"Glow",CSetting::TYPE_BOOL,IDS_GLOW,IDS_GLOW_TIP,0,CSetting::FLAG_WARM,L"ShowCaption"},
	{L"GlowColor",CSetting::TYPE_COLOR,IDS_GLOW_COLOR,IDS_GLOW_COLOR_TIP,0xFFFFFF,CSetting::FLAG_WARM|(4<<24),L"Glow"},
	{L"MaxGlow",CSetting::TYPE_BOOL,IDS_MAXGLOW,IDS_MAXGLOW_TIP,0,CSetting::FLAG_WARM,L"ShowCaption"},
	{L"MaxGlowColor",CSetting::TYPE_COLOR,IDS_MAXGLOW_COLOR,IDS_MAXGLOW_COLOR_TIP,0xFFFFFF,CSetting::FLAG_WARM|(5<<24),L"MaxGlow"},

{L"StatusBar",CSetting::TYPE_GROUP,IDS_STATUS_SETTINGS},
	{L"ShowProgress",CSetting::TYPE_BOOL,IDS_SHOW_PROGRESS,IDS_SHOW_PROGRESS_TIP,1,CSetting::FLAG_WARM|CSetting::FLAG_BASIC},
	{L"ShowZone",CSetting::TYPE_BOOL,IDS_SHOW_ZONE,IDS_SHOW_ZONE_TIP,1,CSetting::FLAG_WARM|CSetting::FLAG_BASIC},
	{L"ShowProtected",CSetting::TYPE_BOOL,IDS_SHOW_PROTECTED,IDS_SHOW_PROTECTED_TIP,1,CSetting::FLAG_WARM,L"ShowZone"},

{L"Language",CSetting::TYPE_GROUP,IDS_LANGUAGE_SETTINGS,0,0,0,NULL,GetLanguageSettings()},
	{L"Language",CSetting::TYPE_STRING,0,0,L""},

{NULL}
};

void UpdateSettings( void )
{
	bool bVista=(LOWORD(GetVersion())==0x0006);

	BOOL bComposition;
	if (FAILED(DwmIsCompositionEnabled(&bComposition)))
		bComposition=FALSE;
	UpdateSetting(L"Glow",CComVariant(bComposition?1:0),false);
	UpdateSetting(L"MaxGlow",CComVariant((bComposition && !bVista)?1:0),false);

	// create a dummy window to get a theme
	HWND hwnd=CreateWindow(L"#32770",L"",WS_OVERLAPPEDWINDOW,0,0,0,0,NULL,NULL,NULL,0);
	HTHEME theme=OpenThemeData(hwnd,L"Window");
	if (theme)
	{
		HDC hdc=GetDC(NULL);
		int dpi=GetDeviceCaps(hdc,LOGPIXELSY);
		ReleaseDC(NULL,hdc);

		LOGFONT font;
		GetThemeSysFont(theme,TMT_CAPTIONFONT,&font);
		wchar_t text[256];
		const wchar_t *type=font.lfItalic?L"italic":L"normal";
		if (font.lfWeight==FW_BOLD)
			type=font.lfItalic?L"bold_italic":L"bold";
		Sprintf(text,_countof(text),L"%s, %s, %d",font.lfFaceName,type,(-font.lfHeight*72+dpi/2)/dpi);
		UpdateSetting(L"CaptionFont",CComVariant(text),false);

		int color=GetThemeSysColor(theme,COLOR_CAPTIONTEXT);
		UpdateSetting(L"TextColor",CComVariant(color),false);
		UpdateSetting(L"MaxColor",CComVariant((bVista && bComposition)?0xFFFFFF:color),false);
		if (bVista)
			color=GetThemeSysColor(theme,COLOR_INACTIVECAPTIONTEXT);
		UpdateSetting(L"InactiveColor",CComVariant(color),false);
		UpdateSetting(L"InactiveMaxColor",CComVariant((bVista && bComposition)?0xFFFFFF:color),false);

		CloseThemeData(theme);
	}
	else
	{
		UpdateSetting(L"TextColor",CComVariant(0),false);
		UpdateSetting(L"MaxColor",CComVariant(0),false);
		UpdateSetting(L"InactiveColor",CComVariant(0),false);
		UpdateSetting(L"InactiveMaxColor",CComVariant(0),false);
	}
	DestroyWindow(hwnd);
}

void InitSettings( void )
{
	InitSettings(g_Settings,COMPONENT_IE9);
}

void ClosingSettings( HWND hWnd, int flags, int command )
{
	if (command==IDOK)
	{
		if (flags&CSetting::FLAG_WARM)
		{
			if (FindWindow(L"IEFrame",NULL))
				MessageBox(hWnd,LoadStringEx(IDS_NEW_SETTINGS),LoadStringEx(IDS_APP_TITLE),MB_OK|MB_ICONINFORMATION);
		}
	}
}

CSIE9API void ShowIE9Settings( void )
{
	if (!GetSettingBool(L"EnableSettings"))
		return;
	wchar_t title[100];
	DWORD ver=GetVersionEx(g_Instance);
	if (ver)
		Sprintf(title,_countof(title),LoadStringEx(IDS_SETTINGS_TITLE_VER),ver>>24,(ver>>16)&0xFF,ver&0xFFFF);
	else
		Sprintf(title,_countof(title),LoadStringEx(IDS_SETTINGS_TITLE));
	EditSettings(title,true);
}
