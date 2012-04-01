// Classic Shell (c) 2009-2012, Ivo Beltchev
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
	bool bTwoColumns;
	std::vector<std::pair<int,CString>> Variations;
	struct Option
	{
		CString name;
		CString label;
		CString condition;
		bool value;
		bool value2; // the value when the condition is false
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

	enum
	{
		USER_CENTER=10000,
		USER_CENTER1=10001,
		USER_CENTER2=10002,
	};

	enum TNameAlign
	{
		NAME_CENTER,
		NAME_CENTER1,
		NAME_CENTER2,
		NAME_LEFT,
		NAME_LEFT1,
		NAME_LEFT2,
		NAME_RIGHT,
		NAME_RIGHT1,
		NAME_RIGHT2,
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
	int Main_bitmap_slices_X[9];
	int Main_bitmap_slices_Y[3];
	TOpacity Main_opacity;
	TOpacity Main_opacity2;

	bool Main_large_icons;

	HFONT Main_font;
	HFONT Main_font2;
	int Main_glow_size;
	int Main_glow_size2;
	COLORREF Main_background;
	COLORREF Main_background2;
	COLORREF Main_text_color[4]; // normal, selected, disabled, selected+disabled
	COLORREF Main_text_color2[4]; // normal, selected, disabled, selected+disabled
	COLORREF Main_arrow_color[2]; // normal, selected
	COLORREF Main_arrow_color2[2]; // normal, selected
	RECT Main_padding;
	RECT Main_padding2;
	bool Main_selectionColor;
	bool Main_selectionColor2;
	bool Main_selection32; // 32-bit bitmap
	bool Main_selection232; // 32-bit bitmap
	union
	{
		HBITMAP bmp; // if Main_selectionColor is false
		COLORREF color; // if Main_selectionColor is true
	} Main_selection,  Main_selection2;
	int Main_selection_slices_X[3];
	int Main_selection_slices_Y[3];
	int Main_selection_slices_X2[3];
	int Main_selection_slices_Y2[3];
	HBITMAP Main_arrow;
	bool Main_arrow32;
	SIZE Main_arrow_Size;
	SIZE Main_arrow_padding;
	HBITMAP Main_arrow2;
	bool Main_arrow232;
	SIZE Main_arrow_Size2;
	SIZE Main_arrow_padding2;
	HBITMAP Main_separator;
	int Main_separatorHeight;
	int Main_separator_slices_X[3];
	HBITMAP Main_separator2;
	int Main_separatorHeight2;
	int Main_separator_slices_X2[3];
	HBITMAP Main_separatorV;
	int Main_separatorWidth;
	int Main_separator_slices_Y[3];
	HBITMAP Main_icon_frame;
	int Main_icon_frame_slices_X[3];
	int Main_icon_frame_slices_Y[3];
	POINT Main_icon_frame_offset;
	bool Main_thin_frame;
	bool Main_separator32; // 32-bit bitmap
	bool Main_separator232;
	bool Main_separatorV32;
	bool Main_icon_frame32;
	RECT Main_icon_padding;
	RECT Main_icon_padding2;
	RECT Main_text_padding;
	RECT Main_text_padding2;
	HBITMAP Main_pager;
	bool Main_pager32;
	int Main_pager_slices_X[3];
	int Main_pager_slices_Y[3];
	HBITMAP Main_pager_arrows;
	bool Main_pager_arrows32;
	SIZE Main_pager_arrow_Size;

	HBITMAP User_bitmap;
	POINT User_frame_position;
	POINT User_image_offset;
	int User_image_size;
	int User_image_alpha;

	RECT User_name_position;
	TNameAlign User_name_align;
	HFONT User_font;
	COLORREF User_text_color;
	COLORREF User_glow_color;
	int User_glow_size;

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
	int AllPrograms_offset;
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
	SIZE Submenu_arrow_padding;
	bool Submenu_thin_frame;
	HBITMAP Submenu_separator;
	bool Submenu_separator32; // 32-bit bitmap
	int Submenu_separatorHeight;
	int Submenu_separator_slices_X[3];

	HBITMAP Submenu_separatorV;
	bool Submenu_separatorV32; // 32-bit bitmap
	int Submenu_separatorWidth;
	int Submenu_separator_slices_Y[3];
	HBITMAP Submenu_icon_frame;
	int Submenu_icon_frame_slices_X[3];
	int Submenu_icon_frame_slices_Y[3];
	POINT Submenu_icon_frame_offset;
	bool Submenu_icon_frame32; // 32-bit bitmap

	RECT Submenu_icon_padding;
	RECT Submenu_text_padding;
	HBITMAP Submenu_pager;
	bool Submenu_pager32;
	int Submenu_pager_slices_X[3];
	int Submenu_pager_slices_Y[3];
	HBITMAP Submenu_pager_arrows;
	bool Submenu_pager_arrows32;
	SIZE Submenu_pager_arrow_Size;

	// SEARCH SECTION
	HBITMAP Search_bitmap;
	bool Search_bitmap32; // 32-bit bitmap

	MenuSkin( void );
	~MenuSkin( void );
	void Reset( void );

	static wchar_t s_SkinError[1024]; // parsing error. must end on \r\n
};

enum
{
	LOADMENU_RESOURCES=1,
	LOADMENU_MAIN=2,
};

bool LoadMenuSkin( const wchar_t *fname, MenuSkin &skin, const wchar_t *variation, const wchar_t *options, unsigned int flags );
void LoadDefaultMenuSkin( MenuSkin &skin, unsigned int flags );
void FreeMenuSkin( MenuSkin &skin );

// Returns the path to the skin files. path must be _MAX_PATH characters
void GetSkinsPath( wchar_t *path );
