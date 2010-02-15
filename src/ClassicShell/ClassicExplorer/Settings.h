// Classic Shell (c) 2009-2010, Ivo Beltchev
// The sources for Classic Shell are distributed under the MIT open source license

#pragma once

#include "ExplorerBand.h"

const int DEFAULT_BUTTONS=0x1FE; // buttons visible by default
const int DEFAULT_ONLY_BUTTON=CBandWindow::ID_SETTINGS; // use this button when all buttons are hidden (there must be at least one visible button)

void ShowSettings( HWND parent );
void ShowSettingsMenu( HWND parent, int x, int y );
