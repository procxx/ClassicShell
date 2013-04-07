// Classic Shell (c) 2009-2013, Ivo Beltchev
// The sources for Classic Shell are distributed under the MIT open source license

#pragma once

enum TStartButtonType
{
	START_BUTTON_CLASSIC,
	START_BUTTON_AERO,
	START_BUTTON_METRO,
	START_BUTTON_CUSTOM,
};

HWND CreateStartButton( int taskbarId, HWND taskBar, HWND rebar, const RECT &rcTask );
void DestroyStartButton( int taskbarId );
void UpdateStartButton( int taskbarId );
void PressStartButton( int taskbarId, bool bPressed );
SIZE GetStartButtonSize( int taskbarId );
bool IsStartButtonSmallIcons( int taskbarId );
bool IsTaskbarSmallIcons( void );
void TaskBarMouseMove( int taskbarId );
