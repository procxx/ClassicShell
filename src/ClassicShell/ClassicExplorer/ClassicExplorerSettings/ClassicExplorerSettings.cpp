#include <windows.h>
#include <commctrl.h>
#include <shlwapi.h>

// Manifest to enable the 6.0 common controls
#pragma comment(linker, \
	"\"/manifestdependency:type='Win32' "\
	"name='Microsoft.Windows.Common-Controls' "\
	"version='6.0.0.0' "\
	"processorArchitecture='*' "\
	"publicKeyToken='6595b64144ccf1df' "\
	"language='*'\"")

// A simple program that loads ClassicExplorer32.dll and calls the ShowExplorerSettings function
// Why not use rundll32 instead? Because it doesn't include the correct manifest for comctl32.dll
int WINAPI wWinMain( HINSTANCE hInstance, HINSTANCE hPrevInstance, LPTSTR lpstrCmdLine, int nCmdShow )
{
	INITCOMMONCONTROLSEX init={sizeof(init),ICC_STANDARD_CLASSES};
	InitCommonControlsEx(&init);

	wchar_t path[_MAX_PATH];
	GetModuleFileName(NULL,path,_countof(path));
	*PathFindFileName(path)=0;
	wcscat_s(path,L"ClassicExplorer32.dll");

	HMODULE dll=LoadLibrary(path);
	if (!dll) return 1;
	
	FARPROC proc=GetProcAddress(dll,"ShowExplorerSettings");
	if (!proc) return 2;
	proc();
	return 0;
}
