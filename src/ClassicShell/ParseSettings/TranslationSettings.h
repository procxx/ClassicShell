// Classic Shell (c) 2009-2010, Ivo Beltchev
// The sources for Classic Shell are distributed under the MIT open source license

#pragma once

// Parses the settings from an ini file. Supports UTF16, UTF8 or ANSI files
// Use forceLang for force a specific language
void ParseTranslations( const wchar_t *fname, const wchar_t *forceLang );

// Returns a setting with the given name. If no setting is found, returns def
const wchar_t *FindTranslation( const char *name, const wchar_t *def );
const wchar_t *FindTranslation( const wchar_t *name, const wchar_t *def );

// Checks for right-to-left languages
bool IsLanguageRTL( void );
