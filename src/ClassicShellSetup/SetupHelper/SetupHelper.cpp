// Classic Shell (c) 2009-2011, Ivo Beltchev
// The sources for Classic Shell are distributed under the MIT open source license

#include <windows.h>
#include <stdio.h>
#include <vector>
#include "StringUtils.h"
#include "FNVHash.h"
#include "SettingsParser.h"
#include "resource.h"

// Manifest to enable the 6.0 common controls
#pragma comment(linker, \
	"\"/manifestdependency:type='Win32' "\
	"name='Microsoft.Windows.Common-Controls' "\
	"version='6.0.0.0' "\
	"processorArchitecture='*' "\
	"publicKeyToken='6595b64144ccf1df' "\
	"language='*'\"")

HINSTANCE g_hInstance;

///////////////////////////////////////////////////////////////////////////////

static unsigned int CalcFileFNV( const wchar_t *fname )
{
	FILE *f=NULL;
	if (_wfopen_s(&f,fname,L"rb") || !f) return 0;
	fseek(f,0,SEEK_END);
	int size=ftell(f);
	fseek(f,0,SEEK_SET);
	std::vector<unsigned char> buf(size);
	if (size==0 || fread(&buf[0],1,size,f)!=size)
	{
		fclose(f);
		return 0;
	}
	fclose(f);
	return CalcFNVHash(&buf[0],size,FNV_HASH0);
}

static void Printf( const char *format, ... )
{
	char buf[1024];
	va_list args;
	va_start(args,format);
	int len=Vsprintf(buf,_countof(buf),format,args);
	va_end(args);
	DWORD q;
	WriteFile(GetStdHandle(STD_OUTPUT_HANDLE),buf,len,&q,NULL);
#if _DEBUG
	OutputDebugStringA(buf);
#endif
}

///////////////////////////////////////////////////////////////////////////////

struct IniFile
{
	const wchar_t *name;
	bool bExplorer;
};

const IniFile g_Inis[]=
{
	{L"ExplorerL10N.ini",true},
	{L"StartMenuL10N.ini",false},
};

int CalcIniChecksum( wchar_t *const *params, int count )
{
	if (count<3) return 2;
	AttachConsole(ATTACH_PARENT_PROCESS);
	wchar_t path[_MAX_PATH];
	unsigned int fnvs[_countof(g_Inis)];
	for (int i=0;i<_countof(g_Inis);i++)
	{
		Sprintf(path,_countof(path),L"%s\\%s",params[g_Inis[i].bExplorer?1:2],g_Inis[i].name);
		fnvs[i]=CalcFileFNV(path);
		if (!fnvs[i])
			Printf("Failed to read %S\n",path);
	}

	// save fnvs
	FILE *f=NULL;
	if (_wfopen_s(&f,L"inichecksum.bin",L"wb") || !f)
	{
		Printf("Failed to open inichecksum.bin\n");
		return 1;
	}
	fwrite(fnvs,4,_countof(fnvs),f);
	fclose(f);
	return 0;
}

///////////////////////////////////////////////////////////////////////////////
int DeleteIniFiles( wchar_t *const *params, int count )
{
	if (count<3) return 2;
	bool bQuiet=(_wtol(params[1])<3);
	SetCurrentDirectory(params[2]);
	if (!bQuiet)
	{
		unsigned int fnvs[_countof(g_Inis)];
		memset(fnvs,0,sizeof(fnvs));
		FILE *f=NULL;
		if (!_wfopen_s(&f,L"inichecksum.bin",L"rb") && f)
		{
			fread(fnvs,4,_countof(fnvs),f);
			fclose(f);
		}

		bool bDiffs[_countof(g_Inis)];
		bool bDiff=false;

		for (int i=0;i<_countof(g_Inis);i++)
		{
			unsigned int fnv=CalcFileFNV(g_Inis[i].name);
			bDiffs[i]=(fnv && fnvs[i]!=fnv);
			if (bDiffs[i])
				bDiff=true;
		}

		if (bDiff)
		{
			wchar_t strTitle[256];
			if (!LoadString(g_hInstance,IDS_APP_TITLE,strTitle,_countof(strTitle))) strTitle[0]=0;
			wchar_t strText[1024];
			if (!LoadString(g_hInstance,IDS_COPY_INI,strText,_countof(strText))) strText[0]=0;

			if (MessageBox(NULL,strText,strTitle,MB_YESNO|MB_SYSTEMMODAL)==IDYES)
			{
				for (int i=0;i<_countof(g_Inis);i++)
				{
					if (!bDiffs[i]) continue;
					wchar_t path[_MAX_PATH];
					Sprintf(path,_countof(path),L"%s.bak",g_Inis[i].name);
					CopyFile(g_Inis[i].name,path,FALSE);
				}
			}
		}
	}

	for (int i=0;i<_countof(g_Inis);i++)
		DeleteFile(g_Inis[i].name);
	return 0;
}

///////////////////////////////////////////////////////////////////////////////

int CalcMsiChecksum( wchar_t *const *params, int count )
{
	if (count<2) return 2;

	AttachConsole(ATTACH_PARENT_PROCESS);
	wchar_t path[_MAX_PATH];
	unsigned int fnvs[2];
	Sprintf(path,_countof(path),L"%s\\ClassicShellSetup32.msi",params[1]);
	fnvs[0]=CalcFileFNV(path);
	if (!fnvs[0])
		Printf("Failed to read %S\n",path);
	Sprintf(path,_countof(path),L"%s\\ClassicShellSetup64.msi",params[1]);
	fnvs[1]=CalcFileFNV(path);
	if (!fnvs[1])
		Printf("Failed to read %S\n",path);

	// save fnvs
	FILE *f=NULL;
	if (_wfopen_s(&f,L"msichecksum.bin",L"wb") || !f)
	{
		Printf("Failed to open msichecksum.bin\n");
		return 1;
	}
	fwrite(fnvs,4,_countof(fnvs),f);
	fclose(f);
	return 0;
}

///////////////////////////////////////////////////////////////////////////////

int ExitStartMenu( void )
{
	HKEY hKey=NULL;
	if (RegCreateKeyEx(HKEY_LOCAL_MACHINE,L"SOFTWARE\\IvoSoft\\ClassicShell",0,NULL,REG_OPTION_NON_VOLATILE,KEY_READ|KEY_QUERY_VALUE|KEY_WOW64_64KEY,NULL,&hKey,NULL)==ERROR_SUCCESS)
	{
		DWORD type=0;
		wchar_t path[_MAX_PATH];
		DWORD size=sizeof(path);
		if (RegQueryValueEx(hKey,L"Path",0,&type,(BYTE*)path,&size)==ERROR_SUCCESS && type==REG_SZ)
		{
			STARTUPINFO startupInfo={sizeof(startupInfo)};
			PROCESS_INFORMATION processInfo;
			memset(&processInfo,0,sizeof(processInfo));
			wcscat_s(path,L"ClassicStartMenu.exe");
			HANDLE h=CreateFile(path,GENERIC_READ,FILE_SHARE_READ|FILE_SHARE_WRITE,NULL,OPEN_EXISTING,FILE_ATTRIBUTE_NORMAL,NULL);
			if (h!=INVALID_HANDLE_VALUE)
			{
				CloseHandle(h);
				wcscat_s(path,L" -exit");
				if (CreateProcess(NULL,path,NULL,NULL,TRUE,0,NULL,NULL,&startupInfo,&processInfo))
				{
					CloseHandle(processInfo.hThread);
					WaitForSingleObject(processInfo.hProcess,5000);
					CloseHandle(processInfo.hProcess);
				}
			}
		}
		RegCloseKey(hKey);
	}
	return 0;
}

///////////////////////////////////////////////////////////////////////////////

int MakeEnglishDll( wchar_t *const *params, int count )
{
	if (count<4) return 2;

	AttachConsole(ATTACH_PARENT_PROCESS);

	HMODULE hEn0=LoadLibraryEx(L"en-US.dll",NULL,LOAD_LIBRARY_AS_DATAFILE|LOAD_LIBRARY_AS_IMAGE_RESOURCE);
	if (!hEn0)
	{
		Printf("Failed to open en-US.dll (err: %d)\n",GetLastError());
		return 1;
	}

	std::vector<char> version;
	{
		HRSRC hResInfo=FindResource(hEn0,MAKEINTRESOURCE(VS_VERSION_INFO),RT_VERSION);
		if (hResInfo)
		{
			HGLOBAL hRes=LoadResource(hEn0,hResInfo);
			void *pRes=LockResource(hRes);
			if (pRes)
			{
				DWORD len=SizeofResource(hEn0,hResInfo);
				if (len>=40+sizeof(VS_FIXEDFILEINFO))
				{
					version.resize(len);
					memcpy(&version[0],pRes,len);
				}
			}
		}
	}
	FreeLibrary(hEn0);
	if (version.empty())
	{
		Printf("Can't find version resource in en-US.dll\n");
		return 1;
	}

	HANDLE hEn=BeginUpdateResource(L"en-US.dll",FALSE);
	if (!hEn)
	{
		Printf("Failed to open en-US.dll (err: %d)\n",GetLastError());
		return 1;
	}

	int res=1;
	HMODULE hExplorer=NULL, hMenu=NULL, hIE9=NULL;
	WORD language=MAKELANGID(LANG_ENGLISH,SUBLANG_ENGLISH_US);

	// get version, strings and dialog from ClassicExplorer32.dll
	hExplorer=LoadLibraryEx(params[1],NULL,LOAD_LIBRARY_AS_DATAFILE|LOAD_LIBRARY_AS_IMAGE_RESOURCE);
	if (!hExplorer)
	{
		Printf("Failed to open %S (err: %d)\n",params[1],GetLastError());
		goto qqq;
	}

	// copy version
	{
		HRSRC hResInfo=FindResource(hExplorer,MAKEINTRESOURCE(VS_VERSION_INFO),RT_VERSION);
		void *pRes=NULL;
		if (hResInfo)
		{
			HGLOBAL hRes=LoadResource(hExplorer,hResInfo);
			pRes=LockResource(hRes);
		}
		if (!pRes)
		{
			Printf("Can't find version resource in %S\n",params[1]);
			goto qqq;
		}
		VS_FIXEDFILEINFO *pVer=(VS_FIXEDFILEINFO*)((char*)pRes+40);
		VS_FIXEDFILEINFO *pVer0=(VS_FIXEDFILEINFO*)(&version[40]);
		pVer0->dwProductVersionMS=pVer->dwProductVersionMS;
		pVer0->dwProductVersionLS=pVer->dwProductVersionLS;
		pVer0->dwFileVersionMS=pVer->dwFileVersionMS;
		pVer0->dwFileVersionLS=pVer->dwFileVersionLS;
		UpdateResource(hEn,RT_VERSION,MAKEINTRESOURCE(VS_VERSION_INFO),language,&version[0],version.size());
	}

	// copy strings
	for (int i=2000;i<3000;i+=16)
	{
		int id=i/16;
		HRSRC hResInfo=FindResource(hExplorer,MAKEINTRESOURCE(id),RT_STRING);
		if (!hResInfo) continue;
		HGLOBAL hRes=LoadResource(hExplorer,hResInfo);
		void *pRes=LockResource(hRes);
		if (!pRes) continue;
		UpdateResource(hEn,RT_STRING,MAKEINTRESOURCE(id),language,pRes,SizeofResource(hExplorer,hResInfo));
	}

	// copy dialogs
	for (int id=2000;id<2010;id++)
	{
		HRSRC hResInfo=FindResource(hExplorer,MAKEINTRESOURCE(id),RT_DIALOG);
		if (!hResInfo) continue;
		HGLOBAL hRes=LoadResource(hExplorer,hResInfo);
		void *pRes=LockResource(hRes);
		if (!pRes) continue;
		UpdateResource(hEn,RT_DIALOG,MAKEINTRESOURCE(id),language,pRes,SizeofResource(hExplorer,hResInfo));
	}

	// get strings and dialog from ClassicStartMenuDLL.dll
	hMenu=LoadLibraryEx(params[2],NULL,LOAD_LIBRARY_AS_DATAFILE|LOAD_LIBRARY_AS_IMAGE_RESOURCE);
	if (!hMenu)
	{
		Printf("Failed to open %S (err: %d)\n",params[2],GetLastError());
		goto qqq;
	}

	// copy strings
	for (int i=3000;i<5000;i+=16)
	{
		int id=i/16;
		HRSRC hResInfo=FindResource(hMenu,MAKEINTRESOURCE(id),RT_STRING);
		if (!hResInfo) continue;
		HGLOBAL hRes=LoadResource(hMenu,hResInfo);
		void *pRes=LockResource(hRes);
		if (!pRes) continue;
		UpdateResource(hEn,RT_STRING,MAKEINTRESOURCE(id),language,pRes,SizeofResource(hMenu,hResInfo));
	}

	// copy dialogs
	for (int id=3000;id<4010;id++)
	{
		HRSRC hResInfo=FindResource(hMenu,MAKEINTRESOURCE(id),RT_DIALOG);
		if (!hResInfo) continue;
		HGLOBAL hRes=LoadResource(hMenu,hResInfo);
		void *pRes=LockResource(hRes);
		if (!pRes) continue;
		UpdateResource(hEn,RT_DIALOG,MAKEINTRESOURCE(id),language,pRes,SizeofResource(hMenu,hResInfo));
	}

	// get strings from ClassicIE9DLL.dll
	hIE9=LoadLibraryEx(params[3],NULL,LOAD_LIBRARY_AS_DATAFILE|LOAD_LIBRARY_AS_IMAGE_RESOURCE);
	if (!hIE9)
	{
		Printf("Failed to open %S (err: %d)\n",params[3],GetLastError());
		goto qqq;
	}

	// copy strings
	for (int i=5000;i<6000;i+=16)
	{
		int id=i/16;
		HRSRC hResInfo=FindResource(hIE9,MAKEINTRESOURCE(id),RT_STRING);
		if (!hResInfo) continue;
		HGLOBAL hRes=LoadResource(hIE9,hResInfo);
		void *pRes=LockResource(hRes);
		if (!pRes) continue;
		UpdateResource(hEn,RT_STRING,MAKEINTRESOURCE(id),language,pRes,SizeofResource(hIE9,hResInfo));
	}

	res=0;
qqq:
	if (!EndUpdateResource(hEn,res!=0) && res==0)
	{
		Printf("Failed to update en-US.dll (err: %d)\n",GetLastError());
		res=1;
	}
	if (hExplorer) FreeLibrary(hExplorer);
	if (hMenu) FreeLibrary(hMenu);
	if (hIE9) FreeLibrary(hIE9);

	return res;
}

///////////////////////////////////////////////////////////////////////////////

struct KeyHeader
{
	DWORD keyCount;
	DWORD keyMax;
	DWORD valCount;
	DWORD valMax;
	DWORD dataMax;
};

void SaveRegKeyRec(HKEY hKey, HANDLE file, DWORD b64 )
{
	KeyHeader header;
	if (!hKey || RegQueryInfoKey(hKey,NULL,NULL,NULL,&header.keyCount,&header.keyMax,NULL,&header.valCount,&header.valMax,&header.dataMax,NULL,NULL)!=ERROR_SUCCESS)
		memset(&header,0,sizeof(header));
	DWORD q;
	WriteFile(file,&header,sizeof(header),&q,NULL);
	if (header.keyCount==0 && header.valCount==0)
		return;

	std::vector<wchar_t> keyName(header.keyMax+1);
	std::vector<wchar_t> valName(header.valMax+1);
	std::vector<BYTE> data(header.dataMax+1);

	for (int i=0;i<(int)header.keyCount;i++)
	{
		DWORD size=header.keyMax+1;
		if (RegEnumKeyEx(hKey,i,&keyName[0],&size,NULL,NULL,NULL,NULL)==ERROR_SUCCESS)
		{
			size=(size+1)*2;
			WriteFile(file,&size,4,&q,NULL);
			WriteFile(file,&keyName[0],size,&q,NULL);
			HKEY hKey2;
			if (RegOpenKeyEx(hKey,&keyName[0],0,KEY_READ|b64,&hKey2)!=ERROR_SUCCESS)
				hKey2=NULL;
			SaveRegKeyRec(hKey2,file,b64);
			if (hKey2) RegCloseKey(hKey2);
		}
		else
		{
			size=0;
			WriteFile(file,&size,4,&q,NULL);
		}
	}

	for (int i=0;i<(int)header.valCount;i++)
	{
		DWORD size=header.valMax+1;
		DWORD size2=header.dataMax;
		DWORD type;
		if (RegEnumValue(hKey,i,&valName[0],&size,NULL,&type,&data[0],&size2)==ERROR_SUCCESS)
		{
			size=(size+1)*2;
			WriteFile(file,&size,4,&q,NULL);
			WriteFile(file,&valName[0],size,&q,NULL);
			WriteFile(file,&type,4,&q,NULL);
			WriteFile(file,&size2,4,&q,NULL);
			WriteFile(file,&data[0],size2,&q,NULL);
		}
		else
		{
			size=0;
			WriteFile(file,&size,4,&q,NULL);
		}
	}
}

void SaveRegKey( HKEY hKey, const wchar_t *fname, DWORD b64 )
{
	HANDLE file=CreateFile(fname,GENERIC_WRITE,0,NULL,CREATE_ALWAYS,FILE_ATTRIBUTE_NORMAL,NULL);
	if (file==INVALID_HANDLE_VALUE) return;
	SaveRegKeyRec(hKey,file,b64);
	CloseHandle(file);
}

bool RestoreRegKeyRec( HKEY hKey, HANDLE file, DWORD b64 )
{
	KeyHeader header;
	DWORD q;
	if (!ReadFile(file,&header,sizeof(header),&q,NULL) || q!=sizeof(header)) return false;
	if (header.keyCount==0 && header.valCount==0)
		return true;

	std::vector<wchar_t> keyName(header.keyMax+1);
	std::vector<wchar_t> valName(header.valMax+1);
	std::vector<BYTE> data(header.dataMax+1);

	for (int i=0;i<(int)header.keyCount;i++)
	{
		DWORD size;
		if (!ReadFile(file,&size,4,&q,NULL) || q!=4 || size>(header.keyMax+1)*2)
			return false;
		if (size==0) continue;
		if (!ReadFile(file,&keyName[0],size,&q,NULL) || q!=size)
			return false;
		HKEY hKey2;
		if (RegCreateKeyEx(hKey,&keyName[0],NULL,NULL,0,KEY_WRITE|b64,NULL,&hKey2,NULL)!=ERROR_SUCCESS)
			return false;
		bool res=RestoreRegKeyRec(hKey2,file,b64);
		RegCloseKey(hKey2);
		if (!res) return false;
	}

	for (int i=0;i<(int)header.valCount;i++)
	{
		DWORD size;
		if (!ReadFile(file,&size,4,&q,NULL) || q!=4 || size>(header.valMax+1)*2)
			return false;
		if (size==0) continue;
		if (!ReadFile(file,&valName[0],size,&q,NULL) || q!=size)
			return false;

		DWORD type;
		if (!ReadFile(file,&type,4,&q,NULL) || q!=4)
			return false;

		DWORD size2;
		if (!ReadFile(file,&size2,4,&q,NULL) || q!=4 || size2>header.dataMax)
			return false;
		if (size2>0 && !ReadFile(file,&data[0],size2,&q,NULL) || q!=size2)
			return false;
		if (_wcsicmp(&valName[0],L"Classic Start Menu")!=0)
			RegSetValueEx(hKey,&valName[0],NULL,type,&data[0],size2);
	}
	return true;
}

void RestoreRegKey( HKEY hKey, const wchar_t *fname, DWORD b64 )
{
	HANDLE file=CreateFile(fname,GENERIC_READ,0,NULL,OPEN_EXISTING,FILE_ATTRIBUTE_NORMAL,NULL);
	if (file==INVALID_HANDLE_VALUE) return;
	RestoreRegKeyRec(hKey,file,b64);
	CloseHandle(file);
}

int RestoreRunKey( wchar_t *const *params, int count )
{
	if (count<2) return 2;
	wchar_t path[_MAX_PATH]=L"%LOCALAPPDATA%\\CSRun.bin";
	DoEnvironmentSubst(path,_countof(path));
	BOOL b64=FALSE;
	IsWow64Process(GetCurrentProcess(),&b64);

	if (_wcsicmp(params[1],L"store")==0)
	{
		DeleteFile(path);
		HKEY hKey;
		LONG err=RegOpenKeyEx(HKEY_LOCAL_MACHINE,L"SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run",0,KEY_READ|(b64?KEY_WOW64_64KEY:0),&hKey);
		if (err!=ERROR_SUCCESS)
			return 3;

		SaveRegKey(hKey,path,b64?KEY_WOW64_64KEY:0);
		RegCloseKey(hKey);
	}
	else if (_wcsicmp(params[1],L"restore")==0)
	{
		HKEY hKey;
		DWORD created;
		LONG err=RegCreateKeyEx(HKEY_LOCAL_MACHINE,L"SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run",0,NULL,0,KEY_WRITE|(b64?KEY_WOW64_64KEY:0),NULL,&hKey,&created);
		if (err!=ERROR_SUCCESS)
			return 3;
		if (created==REG_CREATED_NEW_KEY)
		{
			// restore only if the key was deleted
			RestoreRegKey(hKey,path,b64?KEY_WOW64_64KEY:0);
		}
		RegCloseKey(hKey);
	}
	return 0;
}

///////////////////////////////////////////////////////////////////////////////

// Setup Helper - performs custom actions during Classic Shell install/uninstall
// Usage:
//   crc <explorer path> <start menu path> // creates a file with checksum of all ini files
//   crcmsi <msi path> // creates a file with checksum of both msi files
//   ini <install path> level // backs up and deletes the ini files
//   exitSM // exits the start menu if it is running
//   makeEN <explorer dll> <start menu dll> <ie9 dll> // extracts the localization resources and creates a sample en-US.DLL
//   run store|restore // stores or restores the Run registry key (used to work around a bug in the 2.8.1/2 uninstaller)

int WINAPI wWinMain( HINSTANCE hInstance, HINSTANCE hPrevInstance, LPTSTR lpstrCmdLine, int nCmdShow )
{
//	MessageBox(NULL,lpstrCmdLine,L"Command Line",MB_OK|MB_SYSTEMMODAL);

	int count;
	wchar_t *const *params=CommandLineToArgvW(lpstrCmdLine,&count);
	if (!params) return 1;

	g_hInstance=hInstance;

	for (;count>0;count--,params++)
	{
		if (_wcsicmp(params[0],L"crc")==0)
		{
			return CalcIniChecksum(params,count);
		}

		if (_wcsicmp(params[0],L"crcmsi")==0)
		{
			return CalcMsiChecksum(params,count);
		}

		if (_wcsicmp(params[0],L"ini")==0)
		{
			return DeleteIniFiles(params,count);
		}

		if (_wcsicmp(params[0],L"exitSM")==0)
		{
			return ExitStartMenu();
		}

		if (_wcsicmp(params[0],L"makeEN")==0)
		{
			return MakeEnglishDll(params,count);
		}

		if (_wcsicmp(params[0],L"run")==0)
		{
			return RestoreRunKey(params,count);
		}

	}
	return 1;
}
