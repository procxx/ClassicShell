// Classic Shell (c) 2009-2010, Ivo Beltchev
// The sources for Classic Shell are distributed under the MIT open source license

#pragma once

#include <map>

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

	// Retrieves an icon from a shell folder and child ID
	int GetIcon( IShellFolder *pFolder, PIDLIST_ABSOLUTE path, PCUITEMID_CHILD item, bool bLarge );
	// Retrieves an icon from a file and icon index (index>=0 - icon index, index<0 - resource ID)
	int GetIcon( const wchar_t *location, int index, bool bLarge );
	// Retrieves an icon from shell32.dll by resource ID
	int GetStdIcon( int id, bool bLarge );
	// Retrieves an icon for a custom menu item
	int GetCustomIcon( const wchar_t *path, bool bLarge );

	// Must be called when the start menu is about to be unloaded
	void StopPreloading( bool bWait );

	static int GetDPI( void ) { return s_DPI; }

private:
	std::map<unsigned int,int> m_LargeCache;
	std::map<unsigned int,int> m_SmallCache;

	HANDLE m_PreloadThread;
	PIDLIST_ABSOLUTE m_GamesPath;

	void ProcessPreloadedIcons( void );


	enum TLoadingStage
	{
		LOAD_STOPPED, // the loading thread is not running (the critical section is invalid)
		LOAD_STOPPING, // the loading thread is stopping
		LOAD_LOADING, // the loading thread is running
	};

	static TLoadingStage s_LoadingStage;
	static int s_DPI;
	static std::map<unsigned int,HICON> s_PreloadedIcons; // queue of preloaded icons ready to be added to the image list
	static CRITICAL_SECTION s_PreloadSection; // protects all access to m_SmallCache and s_PreloadedIcons

	static DWORD CALLBACK PreloadThread( void *param );
	static void LoadFolderIcons( IShellFolder *pFolder, int level );
};

extern CIconManager g_IconManager;
