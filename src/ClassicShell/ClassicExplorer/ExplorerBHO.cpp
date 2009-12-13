// Classic Shell (c) 2009, Ivo Beltchev
// The sources for Classic Shell are distributed under the MIT open source license

// ExplorerBHO.cpp : Implementation of CExplorerBHO

#include "stdafx.h"
#include "ExplorerBHO.h"
#include <uxtheme.h>

// CExplorerBHO - a browser helper object that implements Alt+Enter for the folder tree

__declspec(thread) HHOOK CExplorerBHO::s_Hook; // one hook per thread

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

static LRESULT CALLBACK SubclassTreeProc( HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData )
{
	if (uMsg==WM_SYSKEYDOWN && wParam==VK_RETURN)
	{
		// Alt+Enter is pressed
		// if this message was for the folder tree, show the properties of the selected item
		DWORD FoldersSettings=CExplorerBHO::FOLDERS_DEFAULT;
		CRegKey regSettings;
		if (regSettings.Open(HKEY_CURRENT_USER,L"Software\\IvoSoft\\ClassicExplorer")==ERROR_SUCCESS)
			regSettings.QueryDWORDValue(L"FoldersSettings",FoldersSettings);

		if ((FoldersSettings&CExplorerBHO::FOLDERS_ALTENTER) && ShowTreeProperties(hWnd))
			return 0;
	}
	if (uMsg==TVM_SETEXTENDEDSTYLE && wParam==(TVS_EX_FADEINOUTEXPANDOS|TVS_EX_AUTOHSCROLL|0x80000000) && lParam==0)
	{
		wParam&=0x7FFFFFFF;

		DWORD FoldersSettings=CExplorerBHO::FOLDERS_DEFAULT;
		CRegKey regSettings;
		if (regSettings.Open(HKEY_CURRENT_USER,L"Software\\IvoSoft\\ClassicExplorer")==ERROR_SUCCESS)
			regSettings.QueryDWORDValue(L"FoldersSettings",FoldersSettings);

		if (!(FoldersSettings&CExplorerBHO::FOLDERS_NOFADE))
			wParam&=~TVS_EX_FADEINOUTEXPANDOS;

		if (FoldersSettings&CExplorerBHO::FOLDERS_CLASSIC)
		{
			SetWindowTheme(hWnd,NULL,NULL);
			DWORD style=GetWindowLong(hWnd,GWL_STYLE);
			style&=~TVS_NOHSCROLL;
			if (FoldersSettings&CExplorerBHO::FOLDERS_SIMPLE)
			{
				style|=TVS_SINGLEEXPAND|TVS_TRACKSELECT;
				style&=~TVS_HASLINES;
			}
			else
			{
				style|=TVS_HASLINES;
				style&=~(TVS_SINGLEEXPAND|TVS_TRACKSELECT);
				wParam|=TVS_EX_FADEINOUTEXPANDOS;
				HIMAGELIST images=TreeView_GetImageList(hWnd,TVSIL_NORMAL);
				int cx, cy;
				ImageList_GetIconSize(images,&cx,&cy);
				TreeView_SetIndent(hWnd,cx+3);
			}
			SetWindowLong(hWnd,GWL_STYLE,style);
		}
		else
		{
			wParam&=~TVS_EX_AUTOHSCROLL;
		}

		if (wParam==0) return 0;
	}
	return DefSubclassProc(hWnd,uMsg,wParam,lParam);
}

LRESULT CALLBACK CExplorerBHO::HookExplorer( int nCode, WPARAM wParam, LPARAM lParam )
{
	if (nCode==HCBT_CREATEWND)
	{
		HWND hWnd=(HWND)wParam;
		CBT_CREATEWND *create=(CBT_CREATEWND*)lParam;
		if (create->lpcs->lpszClass>(LPTSTR)0xFFFF && _wcsicmp(create->lpcs->lpszClass,WC_TREEVIEW)==0)
		{
			SetWindowSubclass(hWnd,SubclassTreeProc,((unsigned int)'Clas'<<16)+'Shel',0);
			PostMessage(hWnd,TVM_SETEXTENDEDSTYLE,TVS_EX_FADEINOUTEXPANDOS|TVS_EX_AUTOHSCROLL|0x80000000,0);
			UnhookWindowsHookEx(s_Hook);
			s_Hook=NULL;
			return 0;
		}
	}
	return CallNextHookEx(NULL,nCode,wParam,lParam);
}

HRESULT STDMETHODCALLTYPE CExplorerBHO::SetSite( IUnknown *pUnkSite )
{
	IObjectWithSiteImpl<CExplorerBHO>::SetSite(pUnkSite);

	if (pUnkSite)
	{
		// hook
		if (!s_Hook)
		{
			s_Hook=SetWindowsHookEx(WH_CBT,HookExplorer,NULL,GetCurrentThreadId());
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

bool ShowTreeProperties( HWND hwndTree )
{
	// find the PIDL of the selected item (combine all child PIDLs from the current item and its parents)
	HTREEITEM hItem=TreeView_GetSelection(hwndTree);
	LPITEMIDLIST pidl=NULL;
	while (hItem)
	{
		TVITEMEX info={TVIF_PARAM,hItem};
		TreeView_GetItem(hwndTree,&info);
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
		hItem=TreeView_GetParent(hwndTree,hItem);
	}
	if (pidl)
	{
		// show properties
		SHELLEXECUTEINFO execute={sizeof(execute),SEE_MASK_IDLIST|SEE_MASK_INVOKEIDLIST,NULL,L"properties"};
		execute.lpIDList=pidl;
		execute.nShow=SW_SHOWNORMAL;
		ShellExecuteEx(&execute);
		ILFree(pidl);
		return true;
	}
	return false;
}
