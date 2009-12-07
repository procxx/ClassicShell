// Classic Shell (c) 2009, Ivo Beltchev
// The sources for Classic Shell are distributed under the MIT open source license

#pragma once

// Contains all settings stored in the registry
struct StartMenuSettings
{
	DWORD ShowFavorites;
	DWORD ShowDocuments;
	DWORD ShowLogOff;
	DWORD ShowUndock;
	DWORD ExpandControlPanel;
	DWORD ExpandNetwork;
	DWORD ExpandPrinters;
	DWORD UseSmallIcons;
	DWORD UseTheme;
	DWORD ScrollMenus;
	DWORD RecentDocuments;
};

// Read the settings from the registry
void ReadSettings( StartMenuSettings &settings );

// Show the UI for editing settings
void EditSettings( void );

// Close the settings box
void CloseSettings( void );
