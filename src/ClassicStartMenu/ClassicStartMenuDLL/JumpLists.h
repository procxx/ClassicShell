// Classic Shell (c) 2009-2012, Ivo Beltchev
// The sources for Classic Shell are distributed under the MIT open source license

#pragma once

#include <atlbase.h>
#include <atlstr.h>
#include <vector>

// Returns the App ID and the target exe for the given shortcut
// appid must be _MAX_PATH characters
bool GetAppInfoForLink( PIDLIST_ABSOLUTE pidl, wchar_t *appid );

// Returns true if the given shortcut has a jumplist (it may be empty)
bool HasJumplist( const wchar_t *appid );

struct CJumpItem
{
	enum Type
	{
		TYPE_UNKNOWN,
		TYPE_ITEM, // IShellItem
		TYPE_LINK, // IShellLink
		TYPE_SEPARATOR,
	};
	Type type;
	unsigned int hash;
	bool bHidden;
	bool bHasArguments;
	CString name;
	CComPtr<IUnknown> pItem;
};

struct CJumpGroup
{
	enum Type
	{
		TYPE_RECENT,
		TYPE_FREQUENT,
		TYPE_TASKS,
		TYPE_CUSTOM,
		TYPE_PINNED,
	};
	
	Type type;
	bool bHidden;
	CString name0;
	CString name;
	std::vector<CJumpItem> items;
};

struct CJumpList
{
	DWORD reserved;
	std::vector<CJumpGroup> groups;

	void Clear( void ) { reserved=0; groups.clear(); }
};

// Returns the jumplist for the given shortcut
bool GetJumplist( const wchar_t *appid, CJumpList &list, int maxCount );

// Executes the given item using the correct application
bool ExecuteJumpItem( const wchar_t *appid, PIDLIST_ABSOLUTE appexe, const CJumpItem &item, HWND hwnd );

// Removes the given item from the jumplist
void RemoveJumpItem( const wchar_t *appid, CJumpList &list, int groupIdx, int itemIdx );

// Pins or unpins the given item from the jumplist
void PinJumpItem( const wchar_t *appid, const CJumpList &list, int groupIdx, int itemIdx, bool bPin );

// Creates the app id resolver object
void CreateAppResolver( void );
