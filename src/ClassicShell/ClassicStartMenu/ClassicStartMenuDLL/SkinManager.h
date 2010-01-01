// Classic Shell (c) 2009-2010, Ivo Beltchev
// The sources for Classic Shell are distributed under the MIT open source license

#pragma once

struct MenuSkin
{
	CString About; // the text to show in the About box
	HICON AboutIcon; // the icon to show in the About box
	bool ForceRTL;

	// MAIN BITMAP SECTION - describes the background of the main menu
	HBITMAP Main_bitmap;
	bool Main_bitmap32; // 32-bit bitmap
	int Main_bitmap_slices_X[6];
	int Main_bitmap_slices_Y[3];

	enum TOpacity
	{
		OPACITY_SOLID,
		OPACITY_REGION,
		OPACITY_ALPHA,
		OPACITY_GLASS,
	};

	TOpacity Main_opacity;
	bool Main_large_icons;

	// CAPTION SECTION - describes the caption portion of the main menu
	HFONT Caption_font;
	COLORREF Caption_text_color;
	COLORREF Caption_glow_color;
	int Caption_glow_size;
	RECT Caption_padding;

	// MENU SECTION - describes the menu portion of the main menu
	HFONT Main_font;
	COLORREF Main_background;
	COLORREF Main_text_color[4]; // normal, selected, disabled, selected+disabled
	RECT Main_padding;
	bool Main_selectionColor;
	bool Main_selection32; // 32-bit bitmap
	union
	{
		HBITMAP bmp; // if Main_selectionColor is false
		COLORREF color; // if Main_selectionColor is true
	} Main_selection;
	int Main_selection_slices_X[3];
	int Main_selection_slices_Y[3];
	HBITMAP Main_separator;
	bool Main_thin_frame;
	bool Main_separator32; // 32-bit bitmap
	int Main_separatorHeight;
	int Main_separator_slices_X[3];

	// SUB-MENU SECTION - describes the menu portion of the sub-menu
	HFONT Submenu_font;
	COLORREF Submenu_background;
	COLORREF Submenu_text_color[4]; // normal, selected, disabled, selected+disabled
	RECT Submenu_padding;
	bool Submenu_selectionColor;
	bool Submenu_selection32; // 32-bit bitmap
	union
	{
		HBITMAP bmp; // if Submenu_selectionColor is false
		COLORREF color; // if Submenu_selectionColor is true
	} Submenu_selection;
	int Submenu_selection_slices_X[3];
	int Submenu_selection_slices_Y[3];
	HBITMAP Submenu_separator;
	bool Submenu_separator32; // 32-bit bitmap
	bool Submenu_thin_frame;
	int Submenu_separatorHeight;
	int Submenu_separator_slices_X[3];

	MenuSkin( void );
	~MenuSkin( void );
	void Reset( void );
};

bool LoadMenuSkin( const wchar_t *fname, MenuSkin &skin );
void LoadDefaultMenuSkin( MenuSkin &skin );
void FreeMenuSkin( MenuSkin &skin );

// Returns the path to the skin files. path must be _MAX_PATH characters
void GetSkinsPath( wchar_t *path );
