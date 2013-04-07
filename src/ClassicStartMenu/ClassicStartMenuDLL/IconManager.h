// Classic Shell (c) 2009-2013, Ivo Beltchev
// The sources for Classic Shell are distributed under the MIT open source license

#pragma once

#include <map>
#include <vector>

// CIconManager - global cache for icons

class CIconManager
{
public:
	~CIconManager( void );

	static int LARGE_ICON_SIZE;
	static int SMALL_ICON_SIZE;
	enum
	{
		ICON_INDEX_DEFAULT=1,
		ICON_INDEX_MASK=  0x07FFF,
		ICON_SHIELD_FLAG= 0x08000, // draw the shield overlay
		ICON_TEMP_FLAG=   0x10000, // use the temp image list (for small icons only)
	};

	enum TIconType
	{
		ICON_TYPE_SMALL,
		ICON_TYPE_LARGE,
		ICON_TYPE_TEMP,
	};

	HIMAGELIST m_LargeIcons;
	HIMAGELIST m_SmallIcons;
	HIMAGELIST m_TempSmallIcons;

	// Initializes the manager. Called from DllMain
	void Init( void );

	static int GetDPI( void ) { return s_DPI; }

	// Retrieves an icon from a shell folder and child ID
	int GetIcon( IShellFolder *pFolder, PCUITEMID_CHILD item, bool bLarge );
	// Retrieves an icon from a file and icon index (index>=0 - icon index, index<0 - resource ID)
	int GetIcon( const wchar_t *location, int index, TIconType type );
	// Retrieves an icon from an image file and a background color
	int GetMetroIcon( const wchar_t *packagePath, const wchar_t *iconPath, DWORD color, bool bLarge );
	// Retrieves an icon for a custom menu item
	int GetCustomIcon( const wchar_t *path, bool bLarge );

	// Must be called when the start menu is about to be unloaded
	void StopLoading( void );

	// Processes the icons that are loaded by the background threads
	void ProcessLoadedIcons( void );

	// Resets the temp icon list
	void ResetTempIcons( void );

	// Marks all icons as unused (also locks s_PostloadSection)
	static void ResetUsedIcons( void ); 
	// Marks the icon as used
	static void AddUsedIcon( int icon );
	// Wakes up the post-loading thread (also unlocks s_PostloadSection)
	static void StartPostLoading( void );

private:
	std::map<unsigned int,int> m_LargeCache;
	std::map<unsigned int,int> m_SmallCache;
	std::map<unsigned int,int> m_TempSmallCache;
	int m_TempSmallCount;

	HANDLE m_PreloadThread;
	HANDLE m_PostloadThread;

	enum TLoadingStage
	{
		LOAD_STOPPED, // the loading thread is not running
		LOAD_STOPPING, // the loading thread is stopping
		LOAD_LOADING, // the loading thread is running
	};

	static volatile TLoadingStage s_LoadingStage;
	static int s_DPI;
	static std::map<unsigned int,HICON> s_PreloadedIcons; // queue of preloaded icons ready to be added to the image list
	static std::map<unsigned int,HBITMAP> s_PreloadedBitmaps; // queue of preloaded bitmaps ready to be added to the image list

	struct IconLocation
	{
		IconLocation( void ) { bUsed=bLoaded=bTemp=false; }
		CString location;
		int index;
		unsigned int key;
		bool bUsed;
		bool bTemp;
		bool bLoaded;
	};

	// the locations of the icons in m_SmallIcons (only valid if bTemp is true)
	static std::vector<IconLocation> s_IconLocations; // only the main thread can add items here

	static CRITICAL_SECTION s_PreloadSection; // protects all access to m_SmallCache, s_PreloadedIcons and s_PreloadedBitmaps
	static CRITICAL_SECTION s_PostloadSection; // protects all access to s_IconLocations
	static HANDLE s_PostloadEvent; // used to wake up the postload thread

	static DWORD CALLBACK PreloadThread( void *param );
	static void LoadFolderIcons( IShellFolder *pFolder, int level );
	static DWORD CALLBACK PostloadThread( void *param );
};

extern CIconManager g_IconManager;
