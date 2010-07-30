// Classic Shell (c) 2009-2010, Ivo Beltchev
// The sources for Classic Shell are distributed under the MIT open source license

#pragma once

#include "ParseSettings.h"

// Parses the settings from an ini file. Supports UTF16, UTF8 or ANSI files
void ParseGlobalSettings( const wchar_t *fname );

// Returns a setting with the given name. If no setting is found, returns def
const wchar_t *FindSetting( const char *name, const wchar_t *def=NULL );
const wchar_t *FindSetting( const wchar_t *name, const wchar_t *def=NULL );

// Returns a boolean setting with the given name. If no setting is found, returns def
bool FindSettingBool( const char *name, bool def );

void ParseGlobalTree( const wchar_t *rootName, std::vector<CSettingsParser::TreeItem> &items );
