// Classic Shell (c) 2009, Ivo Beltchev
// The sources for Classic Shell are distributed under the MIT open source license

#pragma once

// Parses the settings from an ini file. Supports UTF16, UTF8 or ANSI files
void ParseSettings( const wchar_t *fname );

// Frees the resources
void FreeSettings( void );

// Returns a setting with the given name. If no setting is found, returns def
const wchar_t *FindSetting( const char *name, const wchar_t *def );

// Checks for right-to-left languages
bool IsLanguageRTL( void );
