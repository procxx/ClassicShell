// Classic Shell (c) 2009-2012, Ivo Beltchev
// The sources for Classic Shell are distributed under the MIT open source license

#include "stdafx.h"
#include "resource.h"
#include "ClassicIE9DLL_i.h"
#include "ClassicIE9DLL.h"
#include "Settings.h"
#include "dllmain.h"

CSIE9API void LogMessage( const char *text, ... )
{
	if (GetSettingInt(L"LogLevel")==0) return;
	wchar_t fname[_MAX_PATH]=L"%LOCALAPPDATA%\\ClassicIE9Log.txt";
	DoEnvironmentSubst(fname,_countof(fname));
	FILE *f;
	if (_wfopen_s(&f,fname,L"a+b")==0)
	{
		va_list args;
		va_start(args,text);
		vfprintf(f,text,args);
		va_end(args);
		fclose(f);
	}
}

// Used to determine whether the DLL can be unloaded by OLE
STDAPI DllCanUnloadNow(void)
{
	return _AtlModule.DllCanUnloadNow();
}


// Returns a class factory to create an object of the requested type
STDAPI DllGetClassObject(REFCLSID rclsid, REFIID riid, LPVOID* ppv)
{
	return _AtlModule.DllGetClassObject(rclsid, riid, ppv);
}


// DllRegisterServer - Adds entries to the system registry
STDAPI DllRegisterServer(void)
{
	// registers object, typelib and all interfaces in typelib
	return _AtlModule.DllRegisterServer();
}


// DllUnregisterServer - Removes entries from the system registry
STDAPI DllUnregisterServer(void)
{
	return _AtlModule.DllUnregisterServer();
}

// DllInstall - Adds/Removes entries to the system registry per user
//              per machine.	
STDAPI DllInstall(BOOL bInstall, LPCWSTR pszCmdLine)
{
	HRESULT hr = E_FAIL;
	static const wchar_t szUserSwitch[] = L"user";

	if (pszCmdLine != NULL)
	{
		if (_wcsnicmp(pszCmdLine, szUserSwitch, _countof(szUserSwitch)) == 0)
		{
			AtlSetPerUserRegistration(true);
		}
	}

	if (bInstall)
	{	
		hr = DllRegisterServer();
		if (FAILED(hr))
		{	
			DllUnregisterServer();
		}
	}
	else
	{
		hr = DllUnregisterServer();
	}

	return hr;
}
