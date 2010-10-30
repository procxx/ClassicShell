// Classic Shell (c) 2009-2010, Ivo Beltchev
// The sources for Classic Shell are distributed under the MIT open source license

#pragma once

// Parses the settings from an ini file. Supports UTF16, UTF8 or ANSI files
// Use forceLang for force a specific language
void ParseTranslations( const wchar_t *fname, const wchar_t *forceLang );

// Loads text overrides from the given module. They must be in a "L10N" resource with ID=1
void LoadTranslationOverrides( HMODULE hModule );

// Returns a setting with the given name. If no setting is found, returns def
const wchar_t *FindTranslation( const wchar_t *name, const wchar_t *def );

// Checks for right-to-left languages
bool IsLanguageRTL( void );
