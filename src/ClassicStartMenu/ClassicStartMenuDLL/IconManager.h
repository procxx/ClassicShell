// Classic Shell (c) 2009-2012, Ivo Beltchev
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
	HIMAGELIST m_LargeIcons;
	HIMAGELIST m_SmallIcons;

	// Initializes the manager. Called from DllMain
	void Init( void );

	static int GetDPI( void ) { return s_DPI; }

	// Retrieves an icon from a shell folder and child ID
	int GetIcon( IShellFolder *pFolder, PCUITEMID_CHILD item, bool bLarge );
	// Retrieves an icon from a file and icon index (index>=0 - icon index, index<0 - resource ID)
	int GetIcon( const wchar_t *location, int index, bool bLarge );
	// Retrieves an icon for a custom menu item
	int GetCustomIcon( const wchar_t *path, bool bLarge );

	// Must be called when the start menu is about to be unloaded
	void StopLoading( void );

	// Processes the icons that are loaded by the background threads
	void ProcessLoadedIcons( void );

	// Marks all icons as unused (also locks s_PostloadSection)
	static void ResetUsedIcons( void ); 
	// Marks the icon as used
	static void AddUsedIcon( int icon );
	// Wakes up the post-loading thread (also unlocks s_PostloadSection)
	static void StartPostLoading( void );

private:
	std::map<unsigned int,int> m_LargeCache;
	std::map<unsigned int,int> m_SmallCache;

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
