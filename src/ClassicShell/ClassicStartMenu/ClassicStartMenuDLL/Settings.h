// Classic Shell (c) 2009-2010, Ivo Beltchev
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
	DWORD ExpandLinks;
	DWORD ScrollMenus;
	DWORD ConfirmLogOff;
	DWORD RecentDocuments;
	DWORD Hotkey; // 0 - Win key, 1 - no key
	CString SkinName;
};

// Read the settings from the registry
void ReadSettings( StartMenuSettings &settings );

// Show the UI for editing settings
void EditSettings( bool bModal );

// Close the settings box
void CloseSettings( void );

// Process the dialog messages for the settings box
bool IsSettingsMessage( MSG *msg );
