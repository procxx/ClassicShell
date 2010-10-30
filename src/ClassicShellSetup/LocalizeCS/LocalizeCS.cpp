// Classic Shell (c) 2009-2010, Ivo Beltchev
// The sources for Classic Shell are distributed under the MIT open source license

#include <windows.h>
#include <stdio.h>
#include "StringUtils.h"
#include "FNVHash.h"
#include "SettingsParser.h"
#include <vector>
#include <map>

// LocalizeCS.exe is a helper tool that can extract text from a DLL to a CSV, and replace the string table in a DLL with the text from a CSV

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

static void UnsescapeString( wchar_t *string )
{
	wchar_t *dst=string;
	int len=Strlen(string);
	bool bQuoted=false;
	if (string[0]=='"' && string[len-1]=='"')
	{
		bQuoted=true;
		string[len-1]=0;
		if (*string) string++;
	}
	
	for (const wchar_t *src=string;*src;src++)
	{
		if (*src=='\\')
		{
			src++;
			if (!*src) break;
			if (*src=='t')
				*dst++='\t';
			else if (*src=='r')
				*dst++='\r';
			else if (*src=='n')
				*dst++='\n';
			else
				*dst++=*src;
		}
		else if (*src=='"' && bQuoted)
		{
			src++;
			if (!*src) break;
			*dst++=*src;
		}
		else
			*dst++=*src;
	}
	*dst=0;
}

static void WriteString( HANDLE csv, int id, const wchar_t *string, int len, CSettingsParser &parser, int subid=INT_MAX )
{
	DWORD q;
	wchar_t buf[256];
	int len2;
	if (subid!=INT_MAX)
		len2=Sprintf(buf,_countof(buf),L"%d/%d",id,subid&65535);
	else
		len2=Sprintf(buf,_countof(buf),L"%d",id);

	const wchar_t *comment=parser.FindSetting(buf);
	if (comment && _wcsicmp(comment,L"ignore")==0)
		return;

	WriteFile(csv,buf,len2*2,&q,NULL);
	WriteFile(csv,L"\t",2,&q,NULL);

	for (int i=0;i<len;i++)
	{
		WORD c=string[i];
		if (c=='\t')
			WriteFile(csv,L"\\t",4,&q,NULL);
		else if (c=='\r')
			WriteFile(csv,L"\\r",4,&q,NULL);
		else if (c=='\n')
			WriteFile(csv,L"\\n",4,&q,NULL);
		else if (c=='\\')
			WriteFile(csv,L"\\\\",4,&q,NULL);
		else
			WriteFile(csv,&c,2,&q,NULL);
	}

	WriteFile(csv,L"\t\t",4,&q,NULL);
	if (comment)
		WriteFile(csv,comment,Strlen(comment)*2,&q,NULL);
	WriteFile(csv,L"\r\n",4,&q,NULL);
}

static INT_PTR CALLBACK DefaultDlgProc( HWND hwndDlg, UINT uMsg, WPARAM wParam, LPARAM lParam )
{
	if (uMsg==WM_WINDOWPOSCHANGING)
	{
		WINDOWPOS *pos=(WINDOWPOS*)lParam;
		pos->flags&=~SWP_SHOWWINDOW;
	}
	return FALSE;
}

static void WriteDialog( HWND parent, HANDLE csv, int id, HINSTANCE hInstance, const DLGTEMPLATE *pTemplate, CSettingsParser &parser )
{
	HWND dlg=CreateDialogIndirect(hInstance,pTemplate,parent,DefaultDlgProc);
	if (dlg)
	{
		wchar_t text[256];
		GetWindowText(dlg,text,_countof(text));
		if (*text)
			WriteString(csv,id,text,Strlen(text),parser,0);
		for (HWND child=GetWindow(dlg,GW_CHILD);child;child=GetWindow(child,GW_HWNDNEXT))
		{
			GetWindowText(child,text,_countof(text));
			if (*text)
				WriteString(csv,id,text,Strlen(text),parser,(int)GetWindowLong(child,GWL_ID));
		}
		DestroyWindow(dlg);
	}
}

int ExtractStrings( const wchar_t *dllName, const wchar_t *csvName )
{
	HMODULE hDLL=LoadLibraryEx(dllName,NULL,LOAD_LIBRARY_AS_DATAFILE|LOAD_LIBRARY_AS_IMAGE_RESOURCE);
	if (!hDLL)
	{
		Printf("Failed to open %S (err: %d)\n",dllName,GetLastError());
		return 1;
	}

	int res=1;
	CSettingsParser parser;
	parser.LoadText(L"LocComments.txt");
	parser.ParseText();

	HANDLE hCSV=CreateFile(csvName,GENERIC_WRITE,0,NULL,CREATE_ALWAYS,FILE_ATTRIBUTE_NORMAL,NULL);
	if (hCSV==INVALID_HANDLE_VALUE)
	{
		Printf("Failed to write %S\n",csvName);
		CloseHandle(hDLL);
		return 1;
	}

	wchar_t title[]=L"\xFEFFID\tEnglish\tTranslation\tComment\r\n";
	DWORD q;
	WriteFile(hCSV,title,Strlen(title)*2,&q,NULL);

	// copy strings
	for (int i=2000;i<5000;i+=16)
	{
		int id=i/16;
		HRSRC hResInfo=FindResource(hDLL,MAKEINTRESOURCE(id),RT_STRING);
		if (!hResInfo) continue;
		HGLOBAL hRes=LoadResource(hDLL,hResInfo);
		void *pRes=LockResource(hRes);
		if (!pRes) continue;
		const WORD *data=(WORD*)pRes;
		for (int j=0;j<16;j++)
		{
			int len=*data;
			if (len>0)
				WriteString(hCSV,(id-1)*16+j,(const wchar_t*)data+1,len,parser);
			data+=len+1;
		}
	}

	HWND parent=CreateWindow(L"button",NULL,WS_POPUP,0,0,10,10,NULL,NULL,NULL,NULL);

	// copy dialogs
	for (int id=2000;id<4010;id++)
	{
		if (id>=2010 && id<3000) continue;
		if (id>=3010 && id<4000) continue;
		HRSRC hResInfo=FindResource(hDLL,MAKEINTRESOURCE(id),RT_DIALOG);
		if (!hResInfo) continue;
		HGLOBAL hRes=LoadResource(hDLL,hResInfo);
		void *pRes=LockResource(hRes);
		if (!pRes) continue;
		WriteDialog(parent,hCSV,id,hDLL,(DLGTEMPLATE*)pRes,parser);
	}

	// additional strings
	{
		HRSRC hResInfo=FindResource(hDLL,MAKEINTRESOURCE(1),L"L10N");
		if (hResInfo)
		{
			HGLOBAL hRes=LoadResource(hDLL,hResInfo);
			const wchar_t *pRes=(wchar_t*)LockResource(hRes);
			if (pRes)
			{
				int size=SizeofResource(hDLL,hResInfo)/2;
				if (*pRes==L'\xFEFF')
					pRes++, size--;
				wchar_t *pBuf=new wchar_t[size+1];
				memcpy(pBuf,pRes,size*2);
				pBuf[size]=0;
				for (int i=0;i<size;i++)
					if (pBuf[i]=='=')
						pBuf[i]='\t';
				WriteFile(hCSV,pBuf,size*2,&q,NULL);
			}
		}
	}

	CloseHandle(hCSV);
	DestroyWindow(parent);
	if (hDLL) FreeLibrary(hDLL);

	return res;
}

static BOOL CALLBACK EnumResLangProc( HMODULE hModule, LPCTSTR lpszType, LPCTSTR lpszName, WORD wIDLanguage, LONG_PTR lParam )
{
	if (IS_INTRESOURCE(lpszName))
	{
		std::vector<std::pair<int,WORD>> &oldStrings=*(std::vector<std::pair<int,WORD>>*)lParam;
		oldStrings.push_back(std::pair<int,WORD>(PtrToInt(lpszName),wIDLanguage));
	}
	return TRUE;
}

static BOOL CALLBACK EnumResNameProc( HMODULE hModule, LPCTSTR lpszType, LPTSTR lpszName, LONG_PTR lParam )
{
	EnumResourceLanguages(hModule,lpszType,lpszName,EnumResLangProc,lParam);
	return TRUE;
}

static int ImportStrings( const wchar_t *dllName, const wchar_t *csvName )
{
	HANDLE hCSV=CreateFile(csvName,GENERIC_READ,FILE_SHARE_READ,NULL,OPEN_EXISTING,FILE_ATTRIBUTE_NORMAL,NULL);
	if (hCSV==INVALID_HANDLE_VALUE)
	{
		Printf("Failed to read %S\n",csvName);
		return 1;
	}

	int size=SetFilePointer(hCSV,0,NULL,FILE_END)/2;
	SetFilePointer(hCSV,0,NULL,FILE_BEGIN);
	std::vector<wchar_t> buf(size+1);
	DWORD q;
	ReadFile(hCSV,&buf[0],size*2,&q,NULL);
	CloseHandle(hCSV);
	buf[size]=0;
	for (int i=0;i<size;i++)
		if (buf[i]=='\r' || buf[i]=='\n')
			buf[i]=0;

	std::map<int,const wchar_t*> lines;
	wchar_t *str=&buf[0];
	int min=100000, max=0;
	while (str<&buf[size])
	{
		int len=Strlen(str);
		wchar_t *next=str+len+1;
		wchar_t *tab=wcschr(str,'\t');
		if (tab)
		{
			*tab=0;
			int id=_wtol(str);
			bool bNumber=id>0;
			for (;*str;str++)
				if (*str<'0' || *str>'9')
				{
					bNumber=false;
					break;
				}
			if (bNumber)
			{
				tab=wcschr(tab+1,'\t');
				if (tab)
				{
					str=tab+1;
					tab=wcschr(str,'\t');
					if (tab) *tab=0;
					if (lines.find(id)!=lines.end())
					{
						Printf("Duplicate string ID %d\n",id);
						return 1;
					}
					UnsescapeString(str);
					lines[id]=str;
					if (min>id) min=id;
					if (max<id) max=id;
				}
			}
		}
		str=next;
	}

	HMODULE hDll=LoadLibraryEx(dllName,NULL,LOAD_LIBRARY_AS_DATAFILE|LOAD_LIBRARY_AS_IMAGE_RESOURCE);
	if (!hDll)
	{
		Printf("Failed to open %S (err: %d)\n",dllName,GetLastError());
		return 1;
	}

	std::vector<std::pair<int,WORD>> oldStrings;
	EnumResourceNames(hDll,RT_STRING,EnumResNameProc,(LONG_PTR)&oldStrings);
	FreeLibrary(hDll);

	HANDLE hUpdate=BeginUpdateResource(dllName,FALSE);
	if (!hUpdate)
	{
		Printf("Failed to open %S (err: %d)\n",dllName,GetLastError());
		return 1;
	}

	// delete all string resources
	for (int i=0;i<(int)oldStrings.size();i++)
	{
		UpdateResource(hUpdate,RT_STRING,MAKEINTRESOURCE(oldStrings[i].first),oldStrings[i].second,NULL,0);
	}

	// add new string lines
	max+=32;
	for (int i=min;i<max;i+=16)
	{
		int id=i/16;
		int idx=(id-1)*16;
		std::vector<wchar_t> res;
		for (int j=0;j<16;j++,idx++)
		{
			std::map<int,const wchar_t*>::const_iterator it=lines.find(idx);
			const wchar_t *str=L"";
			if (it!=lines.end())
				str=it->second;
			int len=Strlen(str);
			res.push_back((wchar_t)len);
			for (int c=0;c<len;c++)
				res.push_back(str[c]);
		}
		if (res.size()>16)
			UpdateResource(hUpdate,RT_STRING,MAKEINTRESOURCE(id),LANG_NEUTRAL,&res[0],res.size()*2);
	}

	if (!EndUpdateResource(hUpdate,FALSE))
	{
		Printf("Failed to update %S (err: %d)\n",dllName,GetLastError());
		return 1;
	}
	return 0;
}


///////////////////////////////////////////////////////////////////////////////

int WINAPI wWinMain( HINSTANCE hInstance, HINSTANCE hPrevInstance, LPTSTR lpstrCmdLine, int nCmdShow )
{
	AttachConsole(ATTACH_PARENT_PROCESS);

	int argc;
	wchar_t *const *argv=CommandLineToArgvW(lpstrCmdLine,&argc);
	if (!argv) return 1;

	if (argc==3)
	{
		if (_wcsicmp(argv[0],L"extract")==0)
		{
			// extract DLL, CSV
			// extracts the string table, the dialog text, and the L10N text from a DLL and stores it in a CSV
			return ExtractStrings(argv[1],argv[2]);
		}

		if (_wcsicmp(argv[0],L"import")==0)
		{
			// import DLL, CSV
			// replaces the string table in the DLL with the text from the CSV
			return ImportStrings(argv[1],argv[2]);
		}
	}
	return 0;
}
