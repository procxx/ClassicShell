// Classic Shell (c) 2009-2011, Ivo Beltchev
// The sources for Classic Shell are distributed under the MIT open source license

#pragma once

#ifdef CLASSICIE9DLL_EXPORTS
#define CSIE9API __declspec(dllexport)
#else
#define CSIE9API __declspec(dllimport)
#endif

// WH_GETMESSAGE hook for the explorer's GUI thread. The ClassicIE9 exe uses this hook to inject code into the Internet Explorer process
CSIE9API LRESULT CALLBACK HookInject( int code, WPARAM wParam, LPARAM lParam );
CSIE9API void ShowIE9Settings( void );
CSIE9API void LogMessage( const char *text, ... );
