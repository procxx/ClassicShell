// Classic Shell (c) 2009-2010, Ivo Beltchev
// The sources for Classic Shell are distributed under the MIT open source license

#pragma once

#define _ATL_CSTRING_EXPLICIT_CONSTRUCTORS      // some CString constructors will be explicit
#include <atlstr.h>
#include <map>

class CStringSet: public std::map<int,CString>
{
public:
	// Initializes the string database
	void Init( HINSTANCE hInstance );

	// Returns a string by ID (returns "" if the string is missing)
	CString GetString( UINT uID );

private:
	HINSTANCE m_hInstance;

	static BOOL CALLBACK CStringSet::EnumResNameProc( HMODULE hModule, LPCTSTR lpszType, LPTSTR lpszName, LONG_PTR lParam );
};
