// Classic Shell (c) 2009-2010, Ivo Beltchev
// The sources for Classic Shell are distributed under the MIT open source license

// dllmain.h : Declaration of module class.

#include <vector>

class CClassicExplorerModule : public CAtlDllModuleT< CClassicExplorerModule >
{
public :
	DECLARE_LIBID(LIBID_ClassicExplorerLib)
	DECLARE_REGISTRY_APPID_RESOURCEID(IDR_CLASSICEXPLORER, "{65843E27-A491-429F-84A0-30A947E20F92}")
};

extern class CClassicExplorerModule _AtlModule;

// Some utulity functions used by various modules
void ReadIniFile( bool bStartup );
HICON LoadIcon( int iconSize, const wchar_t *path, int index, std::vector<HMODULE> &modules, HMODULE hShell32 );
HICON CreateDisabledIcon( HICON icon, int size );
