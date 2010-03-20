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

// WH_GETMESSAGE hook for the explorer's GUI thread. The start menu exe uses this hook to inject code into the explorer process
STARTMENUAPI LRESULT CALLBACK HookInject( int code, WPARAM wParam, LPARAM lParam );

// Toggle the start menu. bKeyboard - set to true to show the keyboard cues
STARTMENUAPI HWND ToggleStartMenu( HWND startButton, bool bKeyboard );

// Enable or disable the tooltip for the start button
void EnableStartTooltip( bool bEnable );

// Restore the original drop target
void UnhookDropTarget( void );

// Set the hotkeys and controls for the start menu
void SetControls( DWORD hotkeyCSM, DWORD hotkeyNSM, DWORD controls );

extern HWND g_StartButton;
extern HWND g_TaskBar;
extern HWND g_OwnerWindow;
