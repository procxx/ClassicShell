// Classic Shell (c) 2009, Ivo Beltchev
// The sources for Classic Shell are distributed under the MIT open source license

// ExplorerBHO.cpp : Implementation of CExplorerBHO

#include "stdafx.h"
#include "ExplorerBHO.h"


// CExplorerBHO - a browser helper object that implements Alt+Enter for the folder tree

__declspec(thread) HHOOK CExplorerBHO::s_Hook; // one hook per thread
__declspec(thread) HWND CExplorerBHO::s_hwndTree; // one tree control per thread

static BOOL CALLBACK EnumChildProc( HWND hwnd, LPARAM lParam )
{
	wchar_t name[256];
	GetClassName(hwnd,name,_countof(name));
	if (_wcsicmp(name,WC_TREEVIEW)!=0) return TRUE;
	*(HWND*)lParam=hwnd;
	return FALSE;
}

static BOOL CALLBACK FindFolderTreeEnum( HWND hwnd, LPARAM lParam )
{
	wchar_t name[256];
	GetClassName(hwnd,name,_countof(name));
	if (_wcsicmp(name,L"CabinetWClass")!=0) return TRUE;
	EnumChildWindows(hwnd,EnumChildProc,lParam);
	return FALSE;
}

LRESULT CALLBACK CExplorerBHO::HookExplorer( int code, WPARAM wParam, LPARAM lParam )
{
	if (code==HC_ACTION)
	{
		MSG *msg=(MSG*)lParam;
		if (msg->message==WM_SYSKEYDOWN && msg->wParam==VK_RETURN)
		{
			// Alt+Enter is pressed
			if (!s_hwndTree)
			{
				// find the folder tree
				s_hwndTree=(HWND)1;
				EnumThreadWindows(GetCurrentThreadId(),FindFolderTreeEnum,(LPARAM)&s_hwndTree);
			}
			if (msg->hwnd==s_hwndTree)
			{
				// if this message was for the folder tree, show the properties of the selected item
				DWORD EnableAltEnter=1;
				CRegKey regSettings;
				if (regSettings.Open(HKEY_CURRENT_USER,L"Software\\IvoSoft\\ClassicExplorer")==ERROR_SUCCESS)
					regSettings.QueryDWORDValue(L"EnableAltEnter",EnableAltEnter);

				if (EnableAltEnter)
				{
					// find the PIDL of the selected item (combine all child PIDLs from the current item and its parents)
					HTREEITEM hItem=TreeView_GetSelection(s_hwndTree);
					LPITEMIDLIST pidl=NULL;
					while (hItem)
					{
						TVITEMEX info={TVIF_PARAM,hItem};
						TreeView_GetItem(s_hwndTree,&info);
						LPITEMIDLIST **pidl1=(LPITEMIDLIST**)info.lParam;
						if (!pidl1 || !*pidl1 || !**pidl1)
						{
							if (pidl) ILFree(pidl);
							pidl=NULL;
							break;
						}
						LPITEMIDLIST pidl2=pidl?ILCombine(**pidl1,pidl):ILClone(**pidl1);
						if (pidl) ILFree(pidl);
						pidl=pidl2;
						hItem=TreeView_GetParent(s_hwndTree,hItem);
					}
					if (pidl)
					{
						// show properties
						SHELLEXECUTEINFO execute={sizeof(execute),SEE_MASK_IDLIST|SEE_MASK_INVOKEIDLIST,NULL,L"properties"};
						execute.lpIDList=pidl;
						execute.nShow=SW_SHOWNORMAL;
						ShellExecuteEx(&execute);
						ILFree(pidl);
						msg->message=WM_NULL;
					}
				}
			}
		}
	}
	return CallNextHookEx(NULL,code,wParam,lParam);
}

HRESULT STDMETHODCALLTYPE CExplorerBHO::SetSite( IUnknown *pUnkSite )
{
	IObjectWithSiteImpl<CExplorerBHO>::SetSite(pUnkSite);

	if (pUnkSite)
	{
		// hook
		if (!s_Hook)
		{
			s_Hook=SetWindowsHookEx(WH_GETMESSAGE,HookExplorer,NULL,GetCurrentThreadId());
		}
	}
	else
	{
		// unhook
		if (s_Hook)
			UnhookWindowsHookEx(s_Hook);
		s_Hook=NULL;
	}
	return S_OK;
}
