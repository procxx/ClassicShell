// Classic Shell (c) 2009-2010, Ivo Beltchev
// The sources for Classic Shell are distributed under the MIT open source license

#include "TranslationSettings.h"
#include "ParseSettings.h"

static CSettingsParser g_Settings;

// Parses the settings from an ini file. Supports UTF16, UTF8 or ANSI files
void ParseTranslations( const wchar_t *fname )
{
	g_Settings.Reset();

	if (!g_Settings.LoadText(fname)) return;
	g_Settings.ParseText();

	wchar_t languages[100]={0};
	{
		ULONG size=4; // up to 4 languages
		ULONG len=_countof(languages);
		GetThreadPreferredUILanguages(MUI_LANGUAGE_NAME,&size,languages,&len);
		// uncomment this to force a specific language
//		memcpy(languages,L"bg-BG",10); len=7;
		wcscpy_s(languages+len-1,10,L"default");
		languages[len+7]=0;
	}

	g_Settings.FilterLanguages(languages);
}

// Returns a setting with the given name. If no setting is found, returns def
const wchar_t *FindTranslation( const char *name, const wchar_t *def )
{
	return g_Settings.FindSetting(name,def);
}

const wchar_t *FindTranslation( const wchar_t *name, const wchar_t *def )
{
	return g_Settings.FindSetting(name,def);
}

// Checks for right-to-left languages
bool IsLanguageRTL( void )
{
#ifdef _DEBUG
	//	return true; // uncomment this to simulate RTL environment
#endif
	LOCALESIGNATURE localesig;
	LANGID language=GetUserDefaultUILanguage();
	if (GetLocaleInfoW(language,LOCALE_FONTSIGNATURE,(LPWSTR)&localesig,(sizeof(localesig)/sizeof(wchar_t))) && (localesig.lsUsb[3]&0x08000000))
		return true;
	return false;
}
