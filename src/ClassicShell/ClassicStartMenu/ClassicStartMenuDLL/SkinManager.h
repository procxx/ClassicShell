// Classic Shell (c) 2009-2010, Ivo Beltchev
// The sources for Classic Shell are distributed under the MIT open source license

#pragma once

#include <vector>

const int MAX_SKIN_VERSION=2;

struct MenuSkin
{
	CString About; // the text to show in the About box
	HICON AboutIcon; // the icon to show in the About box
	int version; // 1 - skin 1.0 (default), 2 - skin 2.0 (future skins)
	bool ForceRTL;
	std::vector<std::pair<int,CString>> Variations;
	struct Option
	{
		CString name;
		CString label;
		bool value;
	};
	std::vector<Option> Options;

	enum TOpacity
	{
		OPACITY_SOLID,
		OPACITY_REGION,
		OPACITY_ALPHA,
		OPACITY_GLASS,
		OPACITY_FULLALPHA,
		OPACITY_FULLGLASS,
	};

	// CAPTION SECTION - describes the caption portion of the main menu
	HFONT Caption_font;
	COLORREF Caption_text_color;
	COLORREF Caption_glow_color;
	int Caption_glow_size;
	RECT Caption_padding;

	// MENU SECTION - describes the menu portion of the main menu
	HBITMAP Main_bitmap;
	bool Main_bitmap32; // 32-bit bitmap
	int Main_bitmap_slices_X[6];
	int Main_bitmap_slices_Y[3];
	TOpacity Main_opacity;

	bool Main_large_icons;

	HFONT Main_font;
	int Main_glow_size;
	COLORREF Main_background;
	COLORREF Main_text_color[4]; // normal, selected, disabled, selected+disabled
	COLORREF Main_arrow_color[2]; // normal, selected
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
	HBITMAP Main_arrow;
	bool Main_arrow32;
	SIZE Main_arrow_Size;
	HBITMAP Main_separator;
	bool Main_thin_frame;
	bool Main_separator32; // 32-bit bitmap
	int Main_separatorHeight;
	int Main_separator_slices_X[3];
	RECT Main_icon_padding;
	RECT Main_text_padding;
	HBITMAP Main_pager;
	bool Main_pager32;
	int Main_pager_slices_X[3];
	int Main_pager_slices_Y[3];
	HBITMAP Main_pager_arrows;
	bool Main_pager_arrows32;
	SIZE Main_pager_arrow_Size;

	// SUB-MENU SECTION - describes the menu portion of the sub-menu
	HBITMAP Submenu_bitmap;
	bool Submenu_bitmap32; // 32-bit bitmap
	int Submenu_bitmap_slices_X[6];
	int Submenu_bitmap_slices_Y[3];
	TOpacity Submenu_opacity;

	HFONT Submenu_font;
	int Submenu_glow_size;
	COLORREF Submenu_background;
	COLORREF Submenu_text_color[4]; // normal, selected, disabled, selected+disabled
	COLORREF Submenu_arrow_color[2]; // normal, selected
	RECT Submenu_padding;
	int Submenu_offset;
	bool Submenu_selectionColor;
	bool Submenu_selection32; // 32-bit bitmap
	union
	{
		HBITMAP bmp; // if Submenu_selectionColor is false
		COLORREF color; // if Submenu_selectionColor is true
	} Submenu_selection;
	int Submenu_selection_slices_X[3];
	int Submenu_selection_slices_Y[3];
	HBITMAP Submenu_arrow;
	bool Submenu_arrow32;
	SIZE Submenu_arrow_Size;
	bool Submenu_thin_frame;
	HBITMAP Submenu_separator;
	bool Submenu_separator32; // 32-bit bitmap
	int Submenu_separatorHeight;
	int Submenu_separator_slices_X[3];

	HBITMAP Submenu_separatorV;
	bool Submenu_separatorV32; // 32-bit bitmap
	int Submenu_separatorWidth;
	int Submenu_separator_slices_Y[3];

	RECT Submenu_icon_padding;
	RECT Submenu_text_padding;
	HBITMAP Submenu_pager;
	bool Submenu_pager32;
	int Submenu_pager_slices_X[3];
	int Submenu_pager_slices_Y[3];
	HBITMAP Submenu_pager_arrows;
	bool Submenu_pager_arrows32;
	SIZE Submenu_pager_arrow_Size;

	MenuSkin( void );
	~MenuSkin( void );
	void Reset( void );
};

bool LoadMenuSkin( const wchar_t *fname, MenuSkin &skin, const wchar_t *variation, const std::vector<unsigned int> &options, bool bNoResources );
void LoadDefaultMenuSkin( MenuSkin &skin, bool bNoResources );
void FreeMenuSkin( MenuSkin &skin );

// Returns the path to the skin files. path must be _MAX_PATH characters
void GetSkinsPath( wchar_t *path );
