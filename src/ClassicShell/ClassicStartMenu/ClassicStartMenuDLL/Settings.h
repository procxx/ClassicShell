// Classic Shell (c) 2009-2010, Ivo Beltchev
// The sources for Classic Shell are distributed under the MIT open source license

#pragma once

#include <vector>

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
	DWORD HotkeyCSM;
	DWORD HotkeyNSM;
	CString SkinName;
	CString SkinVariation;
	std::vector<unsigned int> SkinOptions; // high 31 bits FNV of the name, 1 bit for the value

	enum
	{
		OPEN_NOTHING,
		OPEN_CLASSIC,
		OPEN_WINDOWS,
	};

	DWORD Controls; // LSB - Click, then Shift+Click, Win, Shift+Win
};

// Read the settings from the registry
void ReadSettings( StartMenuSettings &settings );

// Show the UI for editing settings
void EditSettings( bool bModal );

// Close the settings box
void CloseSettings( void );

// Process the dialog messages for the settings box
bool IsSettingsMessage( MSG *msg );
