// Classic Shell (c) 2009-2010, Ivo Beltchev
// The sources for Classic Shell are distributed under the MIT open source license

#pragma once

#include <vector>

// Loads all strings from hLngInstance
// pDialogs is a NULL-terminated list of dialog IDs. They are loaded from hLngInstance if possible, otherwise from hMainInstance
void LoadTranslationResources( HINSTANCE hMainInstance, HINSTANCE hLngInstance, int *pDialogs );

// Returns a localized string
CString LoadStringEx( int stringID );

// Returns a localized dialog template
DLGTEMPLATE *LoadDialogEx( int dlgID );

// Loads an icon. path can be a path to .ico file, or in the format "module.dll, number"
HICON LoadIcon( int iconSize, const wchar_t *path, std::vector<HMODULE> &modules );

// Loads the icon for the given pidl (file or folder)
HICON LoadIcon( int iconSize, PIDLIST_ABSOLUTE pidl );

// Converts an icon to a bitmap. pBits may be NULL. If bDestroyIcon is true, hIcon will be destroyed
HBITMAP BitmapFromIcon( HICON hIcon, int iconSize, unsigned int **pBits, bool bDestroyIcon );

// Creates a grayscale version of an icon
HICON CreateDisabledIcon( HICON hIcon, int iconSize );

// Returns the version of a given module
DWORD GetVersionEx( HINSTANCE hInstance );
