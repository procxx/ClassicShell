// Classic Shell (c) 2009-2013, Ivo Beltchev
// The sources for Classic Shell are distributed under the MIT open source license

#pragma once

#include <vector>

const int MAX_SKIN_VERSION=2;

struct MenuBitmap
{
	bool bIsBitmap;
	bool bIsOwned; // only valid if bIsBitmap and bitmap
	bool bIs32; // only valid if bIsBitmap and bitmap

	void Init( bool bIsColor=false );
	void Reset( bool bIsColor=false );

	HBITMAP GetBitmap( void ) const { return bIsBitmap?bitmap:NULL; }
	COLORREF GetColor( void ) const { return bIsBitmap?0:color; }

	void operator=( HBITMAP bmp ) { bIsBitmap=true; bitmap=bmp; }
	void operator=( COLORREF col ) { bIsBitmap=false; color=col; }

private:
	union
	{
		COLORREF color;
		HBITMAP bitmap;
	};
};

struct MenuSkin
{
	CString About; // the text to show in the About box
	HICON AboutIcon; // the icon to show in the About box
	int version; // 1 - skin 1.0 (default), 2 - skin 2.0 (future skins)
	bool ForceRTL;
	bool bTwoColumns;
	bool Main_FakeGlass;
	bool Submenu_FakeGlass;
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
	MenuBitmap Main_bitmap;
	int Main_bitmap_slices_X[9];
	int Main_bitmap_slices_Y[3];
	TOpacity Main_opacity;
	TOpacity Main_opacity2;

	bool Main_large_icons;

	HFONT Main_font[2];
	HFONT Main_font2[2];
	int Main_glow_size[2]; // text, separator
	int Main_glow_size2[2]; // text, separator
	COLORREF Main_background;
	COLORREF Main_background2;
	COLORREF Main_text_color[5]; // normal, selected, disabled, selected+disabled, separator
	COLORREF Main_text_color2[5]; // normal, selected, disabled, selected+disabled, separator
	COLORREF Main_arrow_color[2]; // normal, selected
	COLORREF Main_arrow_color2[2]; // normal, selected
	RECT Main_padding;
	RECT Main_padding2;
	MenuBitmap Main_selection, Main_selection2;
	int Main_selection_slices_X[3];
	int Main_selection_slices_Y[3];
	int Main_selection_slices_X2[3];
	int Main_selection_slices_Y2[3];
	MenuBitmap Main_split_selection;
	MenuBitmap Main_split_selection2;
	int Main_split_selection_slices_X[6];
	int Main_split_selection_slices_Y[3];
	int Main_split_selection_slices_X2[6];
	int Main_split_selection_slices_Y2[3];
	MenuBitmap Main_arrow;
	SIZE Main_arrow_Size;
	SIZE Main_arrow_padding;
	MenuBitmap Main_arrow2;
	SIZE Main_arrow_Size2;
	SIZE Main_arrow_padding2;
	MenuBitmap Main_separator;
	int Main_separatorHeight;
	int Main_separator_slices_X[3];
	MenuBitmap Main_separator2;
	int Main_separatorHeight2;
	int Main_separator_slices_X2[3];
	MenuBitmap Main_separatorV;
	int Main_separatorWidth;
	int Main_separator_slices_Y[3];
	MenuBitmap Main_icon_frame;
	int Main_icon_frame_slices_X[3];
	int Main_icon_frame_slices_Y[3];
	POINT Main_icon_frame_offset;
	MenuBitmap Main_icon_frame2;
	int Main_icon_frame_slices_X2[3];
	int Main_icon_frame_slices_Y2[3];
	POINT Main_icon_frame_offset2;
	bool Main_thin_frame;
	bool Main_no_icons2;
	RECT Main_icon_padding;
	RECT Main_icon_padding2;
	RECT Main_text_padding[2]; // text, separator
	RECT Main_text_padding2[2]; // text, separator
	MenuBitmap Main_pager;
	int Main_pager_slices_X[3];
	int Main_pager_slices_Y[3];
	MenuBitmap Main_pager_arrows;
	SIZE Main_pager_arrow_Size;

	MenuBitmap User_bitmap;
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
	MenuBitmap Submenu_bitmap;
	int Submenu_bitmap_slices_X[6];
	int Submenu_bitmap_slices_Y[3];
	TOpacity Submenu_opacity;

	HFONT Submenu_font[2];
	int Submenu_glow_size[2]; // text, separator
	COLORREF Submenu_background;
	COLORREF Submenu_text_color[5]; // normal, selected, disabled, selected+disabled, separator
	COLORREF Submenu_arrow_color[2]; // normal, selected
	RECT Submenu_padding;
	int Submenu_offset;
	int AllPrograms_offset;
	MenuBitmap Submenu_selection;
	int Submenu_selection_slices_X[3];
	int Submenu_selection_slices_Y[3];
	MenuBitmap Submenu_split_selection;
	int Submenu_split_selection_slices_X[6];
	int Submenu_split_selection_slices_Y[3];
	MenuBitmap Submenu_arrow;
	SIZE Submenu_arrow_Size;
	SIZE Submenu_arrow_padding;
	bool Submenu_thin_frame;
	MenuBitmap Submenu_separator;
	int Submenu_separatorHeight;
	int Submenu_separator_slices_X[3];

	MenuBitmap Submenu_separatorV;
	int Submenu_separatorWidth;
	int Submenu_separator_slices_Y[3];
	MenuBitmap Submenu_icon_frame;
	int Submenu_icon_frame_slices_X[3];
	int Submenu_icon_frame_slices_Y[3];
	POINT Submenu_icon_frame_offset;

	RECT Submenu_icon_padding;
	RECT Submenu_text_padding[2]; // text, separator
	MenuBitmap Submenu_pager;
	int Submenu_pager_slices_X[3];
	int Submenu_pager_slices_Y[3];
	MenuBitmap Submenu_pager_arrows;
	SIZE Submenu_pager_arrow_Size;

	MenuBitmap Pin_bitmap;
	SIZE Pin_bitmap_Size;

	// SEARCH SECTION
	MenuBitmap Search_bitmap;

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
