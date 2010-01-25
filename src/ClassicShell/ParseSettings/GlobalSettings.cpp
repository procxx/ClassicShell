// Classic Shell (c) 2009-2010, Ivo Beltchev
// The sources for Classic Shell are distributed under the MIT open source license

#include "TranslationSettings.h"
#include "ParseSettings.h"

static CSettingsParser g_Settings;

// Parses the settings from an ini file. Supports UTF16, UTF8 or ANSI files
void ParseGlobalSettings( const wchar_t *fname )
{
	g_Settings.Reset();

	if (!g_Settings.LoadText(fname)) return;
	g_Settings.ParseText();
}

// Returns a setting with the given name. If no setting is found, returns def
const wchar_t *FindSetting( const char *name, const wchar_t *def )
{
	return g_Settings.FindSetting(name,def);
}

const wchar_t *FindSetting( const wchar_t *name, const wchar_t *def )
{
	return g_Settings.FindSetting(name,def);
}

bool FindSettingBool( const char *name, bool def )
{
	const wchar_t *str=FindSetting(name,NULL);
	if (str) return _wtol(str)!=0;
	return def;
}
