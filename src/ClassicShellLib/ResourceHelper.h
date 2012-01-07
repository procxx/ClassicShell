// Classic Shell (c) 2009-2011, Ivo Beltchev
// The sources for Classic Shell are distributed under the MIT open source license

#pragma once

#include <shobjidl.h>
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

// Extracts icon of a given size from a specified location the way Shell does it
HICON ShExtractIcon( const wchar_t *path, int index, int iconSize );
HICON ShExtractIcon( const char *path, int index, int iconSize );

// Converts an icon to a bitmap. pBits may be NULL. If bDestroyIcon is true, hIcon will be destroyed
HBITMAP BitmapFromIcon( HICON hIcon, int iconSize, unsigned int **pBits, bool bDestroyIcon );

// Premultiplies a DIB section by the alpha channel and a given color
void PremultiplyBitmap( HBITMAP hBitmap, COLORREF rgb );

// Creates a grayscale version of an icon
HICON CreateDisabledIcon( HICON hIcon, int iconSize );

// Returns the version of a given module
DWORD GetVersionEx( HINSTANCE hInstance );

// Wrapper for IShellFolder::ParseDisplayName
HRESULT ShParseDisplayName( wchar_t *pszName, PIDLIST_ABSOLUTE *ppidl, SFGAOF sfgaoIn, SFGAOF *psfgaoOut );

// Separates the arguments from the program
// May return NULL if no arguments are found
const wchar_t *SeparateArguments( const wchar_t *command, wchar_t *program );

// Replaces some common paths with environment variables
void UnExpandEnvStrings( const wchar_t *src, wchar_t *dst, int size );
