// Classic Shell (c) 2009-2010, Ivo Beltchev
// The sources for Classic Shell are distributed under the MIT open source license

#pragma once

#include <windows.h>
#include <vector>

#define _ATL_CSTRING_EXPLICIT_CONSTRUCTORS      // some CString constructors will be explicit

#include <atlbase.h>
#include <atlstr.h>

///////////////////////////////////////////////////////////////////////////////

class CSettingsParser
{
public:
	// Reads a file into m_Text
	bool LoadText( const wchar_t *fname );
	// Reads a text resource into m_Text
	bool LoadText( HMODULE hMod, HRSRC hResInfo );

	// Splits m_Text into m_Lines
	void ParseText( void );

	// Filters the settings that belong to the given language
	// lanugages is a 00-terminated list of language names ordered by priority
	void FilterLanguages( const wchar_t *languages );

	// Returns a setting with the given name. If no setting is found, returns def
	const wchar_t *FindSetting( const char *name, const wchar_t *def=NULL );
	const wchar_t *FindSetting( const wchar_t *name, const wchar_t *def=NULL );

	// Frees all resources
	virtual void Reset( void );

protected:
	std::vector<wchar_t> m_Text;
	std::vector<const wchar_t*> m_Lines;

private:
	const wchar_t *FindSetting( const wchar_t *name, size_t len );
	void LoadText( const unsigned char *buf, int size );
};

///////////////////////////////////////////////////////////////////////////////

class CSkinParser: public CSettingsParser
{
public:
	bool LoadVariation( const wchar_t *fname );
	bool LoadVariation( HMODULE hMod, HRSRC hResInfo );
	virtual void Reset( void );

	// Parses the option from m_Lines[index]. Returns false if index is out of bounds
	bool ParseOption( CString &name, CString &label, bool &value, int index );

	// Filters the conditional groups
	// values/count - list of true options. the rest are assumed to be false
	void FilterConditions( const wchar_t **values, int count );

protected:
	std::vector<wchar_t> m_VarText;
};

///////////////////////////////////////////////////////////////////////////////

const wchar_t *GetToken( const wchar_t *text, wchar_t *token, int size, const wchar_t *separators );
