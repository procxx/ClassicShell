// Classic Shell (c) 2009-2013, Ivo Beltchev
// The sources for Classic Shell are distributed under the MIT open source license

#include "StringSet.h"

static CString CreateString( const WORD *data )
{
	int len=*data;
	data++;
	if (len==0) return NULL;

	CString str;
	wchar_t *ptr=str.GetBuffer(len);
	if (ptr)
	{
		memcpy(ptr,data,len*2);
		ptr[len]=0;
		str.ReleaseBufferSetLength(len);
	}

	return str;
}

BOOL CALLBACK CStringSet::EnumResNameProc( HMODULE hModule, LPCTSTR lpszType, LPTSTR lpszName, LONG_PTR lParam )
{
	CStringSet *set=(CStringSet*)lParam;
	// find resource
	HRSRC hr=FindResource(hModule,lpszName,RT_STRING);
	if (!hr) return TRUE;

	HGLOBAL hg=LoadResource(hModule,hr);
	if (hg)
	{
		const WORD *res=(WORD*)LockResource(hg);
		if (res)
		{
			for (int i=0;i<16;i++)
			{
				UINT id=(((int)lpszName)<<4)+i-16;

				CString str=CreateString(res);
				if (!str.IsEmpty())
					(*set)[id]=str;
				res+=(*res)+1;
			}
			UnlockResource(hg);
		}
	}
	return TRUE;
}

// Initializes the string database
void CStringSet::Init( HINSTANCE hInstance )
{
	m_hInstance=hInstance;
	EnumResourceNames(hInstance,RT_STRING,EnumResNameProc,(LONG_PTR)this);
}

// Returns a string by ID (returns "" if the string is missing)
CString CStringSet::GetString( UINT uID )
{
	// search in the database
	const_iterator it=find(uID);

	if (it!=end())
	{
		if (it->second)
			return it->second;
	}
	return CString();
}
