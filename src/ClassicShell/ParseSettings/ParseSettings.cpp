// Classic Shell (c) 2009-2010, Ivo Beltchev
// The sources for Classic Shell are distributed under the MIT open source license

#include "ParseSettings.h"

// Reads a file into m_Text
bool CSettingsParser::LoadText( const wchar_t *fname )
{
	// read settings file into buf
	FILE *f=NULL;
	if (_wfopen_s(&f,fname,L"rb")) return false;
	if (!f) return false;
	fseek(f,0,SEEK_END);
	int size=ftell(f);
	fseek(f,0,SEEK_SET);
	std::vector<unsigned char> buf(size);
	if (size<4 || fread(&buf[0],1,size,f)!=size)
	{
		fclose(f);
		return false;
	}
	fclose(f);
	LoadText(&buf[0],size);
	return true;
}

// Reads a text resource into m_Text
bool CSettingsParser::LoadText( HMODULE hMod, HRSRC hResInfo )
{
	HGLOBAL hRes=LoadResource(hMod,hResInfo);
	int size=SizeofResource(hMod,hResInfo);
	unsigned char *buf=(unsigned char*)LockResource(hRes);
	if (!buf) return false;
	LoadText(buf,size);
	return true;
}

void CSettingsParser::LoadText( const unsigned char *buf, int size )
{
	// copy buf to text and convert to UTF16
	if (buf[0]==0xFF && buf[1]==0xFE)
	{
		// UTF16
		int len=(size-2)/2;
		m_Text.resize(len+1);
		memcpy(&m_Text[0],&buf[2],size-2);
		m_Text[len]=0;
	}
	else if (buf[0]==0xEF && buf[1]==0xBB && buf[2]==0xBF)
	{
		// UTF8
		int len=MultiByteToWideChar(CP_UTF8,0,(const char*)&buf[3],size-3,NULL,0);
		m_Text.resize(len+1);
		MultiByteToWideChar(CP_UTF8,0,(const char*)&buf[3],size-3,&m_Text[0],len);
		m_Text[len]=0;
	}
	else
	{
		// ACP
		int len=MultiByteToWideChar(CP_ACP,0,(const char*)&buf[0],size,NULL,0);
		m_Text.resize(len+1);
		MultiByteToWideChar(CP_UTF8,0,(const char*)&buf[0],size,&m_Text[0],len);
		m_Text[len]=0;
	}
}

// Splits m_Text into m_Lines
void CSettingsParser::ParseText( void )
{
	// split into lines
	wchar_t *str=&m_Text[0];
	while (*str)
	{
		if (*str!=';') // ignore lines starting with ;
		{
			// trim leading whitespace
			while (*str==' ' || *str=='\t')
				str++;
			m_Lines.push_back(str);
		}
		wchar_t *p1=wcschr(str,'\r');
		wchar_t *p2=wcschr(str,'\n');
		wchar_t *end=&m_Text[m_Text.size()-1];
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
}

// Filters the settings that belong to the given language
// lanugages is a 00-terminated list of language names ordered by priority
void CSettingsParser::FilterLanguages( const wchar_t *languages )
{
	std::vector<const wchar_t*> lines;
	lines.swap(m_Lines);
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
					m_Lines.push_back(line);
				}
				break;
			}
		}
	}
}

// Returns a setting with the given name. If no setting is found, returns def
const wchar_t *CSettingsParser::FindSetting( const char *name, const wchar_t *def )
{
	wchar_t wname[256];
	size_t len=MultiByteToWideChar(1252,0,name,-1,wname,_countof(wname)-1);
	if (len==0 || wname[len-1]!=0)
		return def;
	len--;

	const wchar_t *str=FindSetting(wname,len);
	return str?str:def;
}

const wchar_t *CSettingsParser::FindSetting( const wchar_t *name, const wchar_t *def )
{
	const wchar_t *str=FindSetting(name,wcslen(name));
	return str?str:def;
}

const wchar_t *CSettingsParser::FindSetting( const wchar_t *name, size_t len )
{
	for (size_t i=0;i<m_Lines.size();i++)
	{
		const wchar_t *str=m_Lines[i];
		if (_wcsnicmp(name,str,len)==0)
		{
			str+=len;
			while (*str==' ' || *str=='\t')
				str++;
			if (*str!='=') continue;
			str++;
			while (*str==' ' || *str=='\t')
				str++;
			return *str?str:NULL;
		}
	}

	return NULL;
}

// Frees all resources
void CSettingsParser::Reset( void )
{
	m_Lines.clear();
	m_Text.clear();
}

///////////////////////////////////////////////////////////////////////////////

bool CSkinParser::LoadVariation( const wchar_t *fname )
{
	m_VarText.swap(m_Text);
	bool res=LoadText(fname);
	if (res)
	{
		std::vector<const wchar_t*> lines;
		lines.swap(m_Lines);
		ParseText();
		m_Lines.insert(m_Lines.end(),lines.begin(),lines.end());
	}
	m_VarText.swap(m_Text);
	return res;
}

bool CSkinParser::LoadVariation( HMODULE hMod, HRSRC hResInfo )
{
	m_VarText.swap(m_Text);
	bool res=LoadText(hMod,hResInfo);
	if (res)
	{
		std::vector<const wchar_t*> lines;
		lines.swap(m_Lines);
		ParseText();
		m_Lines.insert(m_Lines.end(),lines.begin(),lines.end());
	}
	m_VarText.swap(m_Text);
	return res;
}

void CSkinParser::Reset( void )
{
	CSettingsParser::Reset();
	m_VarText.clear();
}

///////////////////////////////////////////////////////////////////////////////

const wchar_t *GetToken( const wchar_t *text, wchar_t *token, int size, const wchar_t *separators )
{
	while (*text && wcschr(separators,*text))
		text++;
	const wchar_t *c1=text,*c2;
	if (text[0]=='\"')
	{
		c1++;
		c2=wcschr(c1,'\"');
	}
	else
	{
		c2=c1;
		while (*c2!=0 && !wcschr(separators,*c2))
			c2++;
	}
	if (!c2) c2=text+wcslen(text);
	int l=(int)(c2-c1);
	if (l>size-1) l=size-1;
	memcpy(token,c1,l*2);
	token[l]=0;

	if (*c2) return c2+1;
	else return c2;
}
