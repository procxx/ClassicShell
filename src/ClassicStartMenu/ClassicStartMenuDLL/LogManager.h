// Classic Shell (c) 2009-2011, Ivo Beltchev
// The sources for Classic Shell are distributed under the MIT open source license

#pragma once

// LogManager.h - logging functionality (for debugging)
// Logs different events in the start menu
// Turn it on by setting the LogLevel setting in the registry
// The logging is consuming very little resources when it is turned off

// TLogLevel - Different logging levels. When an event of certain level is generated, it will be logged if the event level
// is not greater than g_LogLevel. The idea is to use higher number for more frequent events.
enum TLogLevel
{
	LOG_NONE=0, // no logging
	LOG_EXECUTE=1, // logs when items are executed
	LOG_OPEN=2, // logs opening and closing of menus
	LOG_MOUSE=3, // logs mouse events (only hovering for now)
};

#define LOG_MENU( LEVEL, TEXT, ... ) if (LEVEL<=g_LogLevel) { LogMessage(TEXT,__VA_ARGS__); }

extern TLogLevel g_LogLevel;
void InitLog( TLogLevel level, const wchar_t *fname );
void CloseLog( void );
void LogMessage( const wchar_t *text, ... );
