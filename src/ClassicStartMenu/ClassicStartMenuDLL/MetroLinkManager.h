// Classic Shell (c) 2009-2013, Ivo Beltchev
// The sources for Classic Shell are distributed under the MIT open source license

#pragma once

#include <atlbase.h>
#include <atlstr.h>
#include <vector>

// Returns a list of all metro links from all packages
void GetMetroLinks( std::vector<CString> &links );

// Returns true if the link with the specified path is a metro app
// Calculates the name and the icon
bool GetMetroLinkInfo( const wchar_t *path, CString *pName, int *pIcon, bool bLarge );

void ExecuteMetroLink( const wchar_t *path );
