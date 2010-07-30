// Classic Shell (c) 2009-2010, Ivo Beltchev
// The sources for Classic Shell are distributed under the MIT open source license

// dllmain.h : Declaration of module
#pragma once

// Some utility functions used by various modules
void ReadIniFile( void );
CString LoadStringEx( int stringID );
HWND CreateSettingsDialog( HWND hWndParent, DLGPROC lpDialogFunc );
DWORD GetVersionEx( HINSTANCE hInstance );
