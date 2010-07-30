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
	// languages is a 00-terminated list of language names ordered by priority
	void FilterLanguages( const wchar_t *languages );

	// Returns a setting with the given name. If no setting is found, returns def
	const wchar_t *FindSetting( const char *name, const wchar_t *def=NULL );
	const wchar_t *FindSetting( const wchar_t *name, const wchar_t *def=NULL );

	// Frees all resources
	virtual void Reset( void );

	struct TreeItem
	{
		CString name; // empty - last child
		int children; // index to the first child. -1 - no children
	};

	// Parses a tree structure of items. The rootName setting must be a list of item names.
	// Then for each name in the list the function will search for name.Items recursively.
	// The last child in the list will have an empty name.
	// Note - the rootName item will not be added to the list
	void ParseTree( const wchar_t *rootName, std::vector<TreeItem> &items );

protected:
	std::vector<wchar_t> m_Text;
	std::vector<const wchar_t*> m_Lines;

private:
	const wchar_t *FindSetting( const wchar_t *name, size_t len );
	void LoadText( const unsigned char *buf, int size );

	int ParseTreeRec( const wchar_t *rootName, std::vector<TreeItem> &items, CString *names, int level );
};

///////////////////////////////////////////////////////////////////////////////

class CSkinParser: public CSettingsParser
{
public:
	bool LoadVariation( const wchar_t *fname );
	bool LoadVariation( HMODULE hMod, HRSRC hResInfo );
	virtual void Reset( void );

	// Parses the option from m_Lines[index]. Returns false if index is out of bounds
	bool ParseOption( CString &name, CString &label, bool &value, CString &condition, bool &value2, int index );

	// Filters the conditional groups
	// values/count - list of true options. the rest are assumed to be false
	void FilterConditions( const wchar_t **values, int count );

protected:
	std::vector<wchar_t> m_VarText;
};

///////////////////////////////////////////////////////////////////////////////

const wchar_t *GetToken( const wchar_t *text, wchar_t *token, int size, const wchar_t *separators );
int EvalCondition( const wchar_t *condition, const wchar_t *const *values, int count );
