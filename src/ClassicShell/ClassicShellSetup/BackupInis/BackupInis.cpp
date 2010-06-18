#include <windows.h>

// Manifest to enable the 6.0 common controls
#pragma comment(linker, \
	"\"/manifestdependency:type='Win32' "\
	"name='Microsoft.Windows.Common-Controls' "\
	"version='6.0.0.0' "\
	"processorArchitecture='*' "\
	"publicKeyToken='6595b64144ccf1df' "\
	"language='*'\"")


int WINAPI wWinMain( HINSTANCE hInstance, HINSTANCE hPrevInstance, LPTSTR lpstrCmdLine, int nCmdShow )
{
	if (MessageBox(NULL,L"Do you want to back up your ini files?\n\nIf you have done modifications to the .ini files, press 'Yes' and they will be renamed to .ini.bak and will not be deleted.",L"Classic Shell Uninstaller",MB_YESNO|MB_SYSTEMMODAL)!=IDYES)
		return 0;
	HKEY hKey;
	if (RegOpenKeyEx(HKEY_LOCAL_MACHINE,L"SOFTWARE\\IvoSoft\\ClassicShell",0,KEY_READ|KEY_WOW64_64KEY,&hKey)!=ERROR_SUCCESS) return 1;
	wchar_t path[_MAX_PATH];
	DWORD size=sizeof(path);
	if (RegQueryValueEx(hKey,L"path",0,NULL,(BYTE*)path,&size)!=ERROR_SUCCESS) return 1;
	RegCloseKey(hKey);

	SetCurrentDirectory(path);

	CopyFile(L"Explorer.ini",L"Explorer.ini.bak",FALSE);
	CopyFile(L"ExplorerL10N.ini",L"ExplorerL10N.ini.bak",FALSE);
	CopyFile(L"StartMenu.ini",L"StartMenu.ini.bak",FALSE);
	CopyFile(L"StartMenuL10N.ini",L"StartMenuL10N.ini.bak",FALSE);
	CopyFile(L"StartMenuItems.ini",L"StartMenuItems.ini.bak",FALSE);

	return 0;
}
