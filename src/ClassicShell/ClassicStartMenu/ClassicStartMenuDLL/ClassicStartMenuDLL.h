// Classic Shell (c) 2009-2010, Ivo Beltchev
// The sources for Classic Shell are distributed under the MIT open source license

#pragma once

#ifdef CLASSICSTARTMENUDLL_EXPORTS
#define STARTMENUAPI __declspec(dllexport)
#else
#define STARTMENUAPI __declspec(dllimport)
#endif

// Find the start button window for the given process
STARTMENUAPI HWND FindStartButton( DWORD process );

// WH_GETMESSAGE hook for the Progman window
STARTMENUAPI LRESULT CALLBACK HookProgMan( int code, WPARAM wParam, LPARAM lParam );

// WH_GETMESSAGE hook for the start button window
STARTMENUAPI LRESULT CALLBACK HookStartButton( int code, WPARAM wParam, LPARAM lParam );

// Toggle the start menu. bKeyboard - set to true to show the keyboard cues
STARTMENUAPI HWND ToggleStartMenu( HWND startButton, bool bKeyboard );

// Enable or disable the tooltip for the start button
void EnableStartTooltip( bool bEnable );

// Restore the original drop target
void UnhookDropTarget( void );

// Set the hotkey for the start menu (0 - Win key, 1 - no hotkey)
void SetHotkey( DWORD hotkey );

extern HWND g_StartButton;
extern HWND g_TaskBar;
extern HWND g_OwnerWindow;
