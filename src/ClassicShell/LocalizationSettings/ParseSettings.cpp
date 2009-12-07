// Classic Shell (c) 2009, Ivo Beltchev
// The sources for Classic Shell are distributed under the MIT open source license

#include "ParseSettings.h"
#include <windows.h>
#include <vector>
#include <string>

static std::vector<std::wstring> g_Settings;

// Parses the settings from an ini file. Supports UTF16, UTF8 or ANSI files
void ParseSettings( const wchar_t *fname )
{
	g_Settings.clear();

	// read settings file into buf
	FILE *f=NULL;
	if (_wfopen_s(&f,fname,L"rb")) return;
	if (!f) return;
	fseek(f,0,SEEK_END);
	int size=ftell(f);
	fseek(f,0,SEEK_SET);
	std::vector<unsigned char> buf(size);
	if (size<4 || fread(&buf[0],1,size,f)!=size)
	{
		fclose(f);
		return;
	}
	fclose(f);

	// copy buf to text and convert to UTF16
	std::vector<wchar_t> text;
	if (buf[0]==0xFF && buf[1]==0xFE)
	{
		// UTF16
		int len=(size-2)/2;
		text.resize(len+1);
		memcpy(&text[0],&buf[2],size-2);
		text[len]=0;
	}
	else if (buf[0]==0xEF && buf[1]==0xBB && buf[2]==0xBF)
	{
		// UTF8
		int len=MultiByteToWideChar(CP_UTF8,0,(const char*)&buf[3],size-3,NULL,0);
		text.resize(len+1);
		MultiByteToWideChar(CP_UTF8,0,(const char*)&buf[3],size-3,&text[0],len);
		text[len]=0;
	}
	else
	{
		// ACP
		int len=MultiByteToWideChar(CP_ACP,0,(const char*)&buf[0],size,NULL,0);
		text.resize(len+1);
		MultiByteToWideChar(CP_UTF8,0,(const char*)&buf[0],size,&text[0],len);
		text[len]=0;
	}

	// split into lines
	std::vector<const wchar_t*> lines;
	wchar_t *str=&text[0];
	while (*str)
	{
		if (*str!=';') // ignore lines starting with ;
		{
			// trim leading whitespace
			while (*str==' ' || *str=='\t')
				str++;
			lines.push_back(str);
		}
		wchar_t *p1=wcschr(str,'\r');
		wchar_t *p2=wcschr(str,'\n');
		wchar_t *end=&text[text.size()-1];
		if (p1) end=p1;
		if (p2 && p2<end) end=p2;

		wchar_t *next=end;
		while (*next=='\r' || *next=='\n')
			next++;

		// trim trailing whitespace
		while (end>str && (*end==' ' || *end=='\t'))
			end--;
		*end=0;
		str=next;
	}

	wchar_t languages[100]={0};
	{
		ULONG size=4; // up to 4 languages
		ULONG len=_countof(languages);
		GetThreadPreferredUILanguages(MUI_LANGUAGE_NAME,&size,languages,&len);
		// uncomment this to force a specific language
//		memcpy(languages,L"ar-SA",10);
//		len=7;
		wcscpy_s(languages+len-1,10,L"default");
		languages[len+7]=0;
	}

	// collect 
	for (const wchar_t *lang=languages;*lang;lang+=wcslen(lang)+1)
	{
		size_t langLen=wcslen(lang);
		for (size_t i=0;i<lines.size();i++)
		{
			const wchar_t *line=lines[i];
			if (*line=='[' && _wcsnicmp(line+1,lang,langLen)==0 && line[langLen+1]==']')
			{
				for (i++;i<lines.size();i++)
				{
					line=lines[i];
					if (*line=='[') break;
					g_Settings.push_back(line);
				}
				break;
			}
		}
	}
}

// Frees the resources
void FreeSettings( void )
{
	g_Settings.clear();
}

// Returns a setting with the given name. If no setting is found, returns def
const wchar_t *FindSetting( const char *name, const wchar_t *def )
{
	wchar_t wname[256];
	size_t len=MultiByteToWideChar(1252,0,name,-1,wname,_countof(wname)-1);
	if (len==0 || wname[len-1]!=0)
		return def;
	len--;

	for (size_t i=0;i<g_Settings.size();i++)
	{
		const wchar_t *str=g_Settings[i].c_str();
		if (_wcsnicmp(wname,str,len)==0)
		{
			str+=len;
			while (*str==' ' || *str=='\t')
				str++;
			if (*str!='=') continue;
			str++;
				while (*str==' ' || *str=='\t')
					str++;
			return str;
		}
	}

	return def;
}

// Checks for right-to-left languages
bool IsLanguageRTL( void )
{
	LOCALESIGNATURE localesig;
	LANGID language=GetUserDefaultUILanguage();
	if (GetLocaleInfoW(language,LOCALE_FONTSIGNATURE,(LPWSTR)&localesig,(sizeof(localesig)/sizeof(wchar_t))) && (localesig.lsUsb[3]&0x08000000))
		return true;
	return false;
}
