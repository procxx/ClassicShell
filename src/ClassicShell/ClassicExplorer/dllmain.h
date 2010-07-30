// Classic Shell (c) 2009-2010, Ivo Beltchev
// The sources for Classic Shell are distributed under the MIT open source license

// dllmain.h : Declaration of module class.
#pragma once

#include "ClassicExplorer_i.h"
#include <vector>

class CClassicExplorerModule : public CAtlDllModuleT< CClassicExplorerModule >
{
public :
	DECLARE_LIBID(LIBID_ClassicExplorerLib)
	DECLARE_REGISTRY_APPID_RESOURCEID(IDR_CLASSICEXPLORER, "{65843E27-A491-429F-84A0-30A947E20F92}")
};

extern class CClassicExplorerModule _AtlModule;

// Some utility functions used by various modules
void ReadIniFile( bool bStartup );
HICON LoadIcon( int iconSize, const wchar_t *path, int index, std::vector<HMODULE> &modules, HMODULE hShell32 );
HBITMAP BitmapFromIcon( HICON icon, int iconSize, unsigned int **pBits, bool bDestroyIcon );
HICON CreateDisabledIcon( HICON icon, int iconSize );
HWND FindChildWindow( HWND hwnd, const wchar_t *className );
CString LoadStringEx( int stringID );
INT_PTR RunSettingsDialog( HWND hWndParent, DLGPROC lpDialogFunc );
DWORD GetVersionEx( HINSTANCE hInstance );

struct TlsData
{
	// one hook for each BHO thread
	HHOOK bhoHook;

	// one hook for each copy thread
	HHOOK copyHook;

	// bCopyMultiFile is true if the first dialog in this thread is multi-file (IDD_FILEMULTI)
	// if so, all the rest are multi-file. this makes the UI consistent (like the position of the Yes button doesn't change)
	bool bCopyMultiFile;
};

TlsData *GetTlsData( void );

enum
{
	FILEUI_FILE=1,
	FILEUI_FOLDER=2,
	FILEUI_MORE=4,
	FILEUI_OTHERAPPS=256, // not only for Explorer

	FILEUI_ALL=FILEUI_FILE|FILEUI_FOLDER|FILEUI_MORE,
	FILEUI_DEFAULT=FILEUI_FILE|FILEUI_FOLDER
};
