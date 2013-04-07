// Classic Shell (c) 2009-2013, Ivo Beltchev
// The sources for Classic Shell are distributed under the MIT open source license

// LogManager.cpp - logging functionality (for debugging)

#include "stdafx.h"
#include "LogManager.h"

TLogLevel g_LogLevel;
static FILE *g_LogFile;
static int g_LogTime;

void InitLog( TLogLevel level, const wchar_t *fname )
{
	CloseLog();
	if (level<=LOG_NONE) return;
	if (_wfopen_s(&g_LogFile,fname,L"wb")==0)
	{
		wchar_t bom=0xFEFF;
		fwrite(&bom,2,1,g_LogFile);
		g_LogLevel=level;
		g_LogTime=GetTickCount();
	}
}

void CloseLog( void )
{
	if (g_LogFile) fclose(g_LogFile);
	g_LogFile=NULL;
	g_LogLevel=LOG_NONE;
}

void LogMessage( const wchar_t *text, ... )
{
	if (!g_LogFile) return;

	wchar_t buf[2048];
	int len=Sprintf(buf,_countof(buf),L"%8d: ",GetTickCount()-g_LogTime);
	fwrite(buf,2,len,g_LogFile);

	va_list args;
	va_start(args,text);
	len=Vsprintf(buf,_countof(buf),text,args);
	va_end(args);
	fwrite(buf,2,len,g_LogFile);

	fwrite(L"\r\n",2,2,g_LogFile);

	fflush(g_LogFile);
}
