// Classic Shell (c) 2009-2012, Ivo Beltchev
// The sources for Classic Shell are distributed under the MIT open source license

#pragma once

enum TStartButtonType
{
	START_BUTTON_CLASSIC,
	START_BUTTON_AERO,
	START_BUTTON_METRO,
	START_BUTTON_CUSTOM,
};

HWND CreateStartButton( HWND taskBar, HWND rebar, const RECT &rcTask );
void DestroyStartButton( void );
void UpdateStartButton( void );
void PressStartButton( bool bPressed );
SIZE GetStartButtonSize( void );
bool IsStartButtonSmallIcons( void );
bool IsTaskbarSmallIcons( void );
void TaskBarMouseMove( void );
