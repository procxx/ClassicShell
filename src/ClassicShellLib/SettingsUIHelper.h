// Classic Shell (c) 2009-2012, Ivo Beltchev
// The sources for Classic Shell are distributed under the MIT open source license

#pragma once

#include "SettingsParser.h"
#include "resource.h"
#include <vector>

class CCommandsTree;
class CSettingsTree;

///////////////////////////////////////////////////////////////////////////////

// CResizeableDlg - a dialog that rearranges its controls when it gets resized
template<class T> class CResizeableDlg: public CDialogImpl<T>
{
public:
	void Create( HWND hWndParent )
	{
		CDialogImpl<T>::Create(hWndParent);
	}

	void Create( HWND hWndParent, DLGTEMPLATE *pTemplate )
	{
		ATLASSUME(m_hWnd == NULL);
		if (!m_thunk.Init(NULL,NULL))
		{
			SetLastError(ERROR_OUTOFMEMORY);
			return;
		}

		_AtlWinModule.AddCreateWndData(&m_thunk.cd,(CDialogImplBaseT<CWindow>*)this);
		HWND hWnd=::CreateDialogIndirect(_AtlBaseModule.GetResourceInstance(),pTemplate,hWndParent,T::StartDialogProc);
		ATLASSUME(m_hWnd==hWnd);
	}

protected:
	enum
	{
		MOVE_LEFT=1,
		MOVE_LEFT2=2,
		MOVE_RIGHT=4,
		MOVE_RIGHT2=8,
		MOVE_TOP=16,
		MOVE_TOP2=32,
		MOVE_BOTTOM=64,
		MOVE_BOTTOM2=128,

		MOVE_MOVE_X=MOVE_LEFT|MOVE_RIGHT,
		MOVE_MOVE_Y=MOVE_TOP|MOVE_BOTTOM,
		MOVE_SIZE_X=MOVE_RIGHT,
		MOVE_SIZE_Y=MOVE_BOTTOM,
		MOVE_LEFT_HALF=MOVE_RIGHT2,
		MOVE_RIGHT_HALF=MOVE_LEFT2|MOVE_RIGHT,
		MOVE_CENTER=MOVE_LEFT2|MOVE_RIGHT2,
		MOVE_TOP_HALF=MOVE_BOTTOM2,
		MOVE_BOTTOM_HALF=MOVE_TOP2|MOVE_BOTTOM,
		MOVE_VCENTER=MOVE_TOP2|MOVE_BOTTOM2,

		MOVE_HORIZONTAL=1,
		MOVE_VERTICAL=2,
		MOVE_GRIPPER=4,
		MOVE_REINITIALIZE=8, // InitResize is called for a second time to recapture the control sizes

		MOVE_MODAL=MOVE_HORIZONTAL|MOVE_VERTICAL|MOVE_GRIPPER,
	};

	struct Control
	{
		int id;
		unsigned int flags;
		HWND hwnd;
		RECT rect0;
	};

	void InitResize( int flags=MOVE_HORIZONTAL|MOVE_VERTICAL )
	{
		m_Flags=flags;
		T *pThis=static_cast<T*>(this);

		int count=0;
		for (const Control *pControl=pThis->GetResizeControls();pControl->id;pControl++)
			count++;
		m_Controls.resize(count);
		if (count>0)
			memcpy(&m_Controls[0],pThis->GetResizeControls(),count*sizeof(Control));

		RECT rc;
		pThis->GetClientRect(&rc);
		if (!(m_Flags&MOVE_REINITIALIZE))
		{
			m_Gripper.m_hWnd=NULL;
			if (m_Flags&MOVE_GRIPPER)
				m_Gripper.Create(L"SCROLLBAR",pThis->m_hWnd,rc,NULL,WS_CHILD|WS_VISIBLE|WS_CLIPSIBLINGS|SBS_SIZEBOX|SBS_SIZEGRIP|SBS_SIZEBOXBOTTOMRIGHTALIGN);
		}
		m_ClientSize.cx=rc.right;
		m_ClientSize.cy=rc.bottom;
		pThis->GetWindowRect(&rc);
		m_WindowSize.cx=rc.right-rc.left;
		m_WindowSize.cy=rc.bottom-rc.top;
		for (std::vector<Control>::iterator it=m_Controls.begin();it!=m_Controls.end();++it)
		{
			it->hwnd=pThis->GetDlgItem(it->id);
			ATLASSERT(it->hwnd);
			if (!it->hwnd) continue;
			::GetWindowRect(it->hwnd,&it->rect0);
			::MapWindowPoints(NULL,m_hWnd,(POINT*)&it->rect0,2);
		}
	}

	void OnSize( void )
	{
		T *pThis=static_cast<T*>(this);
		RECT rc;
		pThis->GetClientRect(&rc);
		int dx=rc.right-m_ClientSize.cx;
		int dy=rc.bottom-m_ClientSize.cy;
		int dx2=dx/2;
		int dy2=dy/2;
		for (std::vector<Control>::iterator it=m_Controls.begin();it!=m_Controls.end();++it)
		{
			if (!it->hwnd) continue;
			int x1=it->rect0.left;
			int y1=it->rect0.top;
			int x2=it->rect0.right;
			int y2=it->rect0.bottom;
			if (it->flags&MOVE_LEFT) x1+=dx;
			else if (it->flags&MOVE_LEFT2) x1+=dx2;
			if (it->flags&MOVE_TOP) y1+=dy;
			else if (it->flags&MOVE_TOP2) y1+=dy2;
			if (it->flags&MOVE_RIGHT) x2+=dx;
			else if (it->flags&MOVE_RIGHT2) x2+=dx2;
			if (it->flags&MOVE_BOTTOM) y2+=dy;
			else if (it->flags&MOVE_BOTTOM2) y2+=dy2;
			::SetWindowPos(it->hwnd,NULL,x1,y1,x2-x1,y2-y1,SWP_NOZORDER|SWP_NOCOPYBITS);
		}
		if (m_Gripper.m_hWnd)
		{
			RECT rc2;
			m_Gripper.GetWindowRect(&rc2);
			int w=rc2.right-rc2.left;
			int h=rc2.bottom-rc2.top;
			m_Gripper.SetWindowPos(HWND_BOTTOM,rc.right-w,rc.bottom-h,w,h,0);
		}
	}

	LRESULT OnGetMinMaxInfo( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled )
	{
		MINMAXINFO *pInfo=(MINMAXINFO*)lParam;
		pInfo->ptMinTrackSize.x=m_WindowSize.cx;
		pInfo->ptMinTrackSize.y=m_WindowSize.cy;
		if (!(m_Flags&MOVE_HORIZONTAL))
			pInfo->ptMaxTrackSize.x=pInfo->ptMinTrackSize.x;
		if (!(m_Flags&MOVE_VERTICAL))
			pInfo->ptMaxTrackSize.y=pInfo->ptMinTrackSize.y;
		return 0;
	}

	void GetStoreRect( RECT &rc )
	{
		GetWindowRect(&rc);
		rc.right-=rc.left+m_WindowSize.cx;
		rc.bottom-=rc.top+m_WindowSize.cy;
	}

	void SetStoreRect( const RECT &rc )
	{
		SetWindowPos(NULL,rc.left,rc.top,m_WindowSize.cx+rc.right,m_WindowSize.cy+rc.bottom,SWP_NOZORDER|SWP_NOCOPYBITS);
		SendMessage(DM_REPOSITION);
	}

private:
	SIZE m_ClientSize;
	SIZE m_WindowSize;
	int m_Flags;
	CWindow m_Gripper;
	std::vector<Control> m_Controls;
};

#define BEGIN_RESIZE_MAP const Control *GetResizeControls( void ) { static Control controls[]={
#define RESIZE_CONTROL(id,flags) {id,flags},
#define END_RESIZE_MAP {0,0}}; return controls; }

///////////////////////////////////////////////////////////////////////////////

struct CStdCommand
{
	const wchar_t *name;
	const wchar_t *displayName;
	int tipID;
	const wchar_t *itemName; // NULL for separator items
	const wchar_t *label;
	const wchar_t *tip;
	const wchar_t *icon; // NULL - force no icon (use L"" for default icon)
	const KNOWNFOLDERID *knownFolder;
	const wchar_t *iconD;
	unsigned int settings;
};

struct CTreeItem
{
	CString name;
	CString command;
	CString link;
	CString label;
	CString tip;
	CString icon;
	CString iconD;
	unsigned int settings;
	const CStdCommand *pStdCommand;
	bool bSeparator;

	CTreeItem( void ) { settings=0; pStdCommand=NULL; bSeparator=false; }
	void SetCommand( CString command, const CStdCommand *pStdCommands );
	unsigned int GetIconKey( void ) const;
	HICON LoadIcon( bool bSmall, std::vector<HMODULE> &modules ) const;
	unsigned int GetIconDKey( unsigned int iconKey ) const;
	HICON LoadIconD( HICON hIcon, std::vector<HMODULE> &modules ) const; // always large
};

///////////////////////////////////////////////////////////////////////////////

const HICON HICON_NONE=(HICON)-1;

class CCustomTreeDlg: public CResizeableDlg<CCustomTreeDlg>
{
public:
	CCustomTreeDlg( bool bMenu, const CStdCommand *pStdCommands );
	~CCustomTreeDlg( void );

	BEGIN_MSG_MAP( CCustomTreeDlg )
		MESSAGE_HANDLER( WM_INITDIALOG, OnInitDialog )
		MESSAGE_HANDLER( WM_SIZE, OnSize )
		MESSAGE_HANDLER( WM_CONTEXTMENU, OnContextMenu )

		NOTIFY_HANDLER( IDC_TREECOMMANDS, TVN_GETINFOTIP, OnGetInfoTip )
		NOTIFY_HANDLER( IDC_TREECOMMANDS, TVN_BEGINDRAG, OnBeginDrag )
		NOTIFY_HANDLER( IDC_TREECOMMANDS, NM_DBLCLK, OnAddItem )
		NOTIFY_HANDLER( IDC_TREECOMMANDS, TVN_KEYDOWN, OnAddItem )
		NOTIFY_HANDLER( IDC_TREEITEMS, TVN_GETINFOTIP, OnGetInfoTip )
		NOTIFY_HANDLER( IDC_TREEITEMS, NM_DBLCLK, OnEditItem )
		NOTIFY_HANDLER( IDC_TREEITEMS, TVN_KEYDOWN, OnEditItem )
		NOTIFY_HANDLER( IDC_TREEITEMS, NM_CUSTOMDRAW, OnCustomDraw )
		REFLECT_NOTIFICATIONS()
	END_MSG_MAP()

	BEGIN_RESIZE_MAP
		RESIZE_CONTROL(IDC_TREEITEMS,MOVE_LEFT_HALF|MOVE_SIZE_Y)
		RESIZE_CONTROL(IDC_STATICMIDDLE,MOVE_CENTER|MOVE_VCENTER)
		RESIZE_CONTROL(IDC_STATICRIGHT,MOVE_CENTER)
		RESIZE_CONTROL(IDC_TREECOMMANDS,MOVE_RIGHT_HALF|MOVE_SIZE_Y)
		RESIZE_CONTROL(IDC_STATICHINT,MOVE_SIZE_X|MOVE_MOVE_Y)
	END_RESIZE_MAP

	void SetGroup( CSetting *pGroup, bool bReset );
	bool IsSeparator( const wchar_t *name );
	void SerializeData( void );

protected:
	// Handler prototypes:
	//  LRESULT MessageHandler(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
	//  LRESULT CommandHandler(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled);
	//  LRESULT NotifyHandler(int idCtrl, LPNMHDR pnmh, BOOL& bHandled);
	LRESULT OnInitDialog( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled );
	LRESULT OnSize( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled );
	LRESULT OnContextMenu( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled );
	LRESULT OnGetInfoTip( int idCtrl, LPNMHDR pnmh, BOOL& bHandled );
	LRESULT OnBeginDrag( int idCtrl, LPNMHDR pnmh, BOOL& bHandled );
	LRESULT OnAddItem( int idCtrl, LPNMHDR pnmh, BOOL& bHandled );
	LRESULT OnEditItem( int idCtrl, LPNMHDR pnmh, BOOL& bHandled );
	LRESULT OnCustomDraw( int idCtrl, LPNMHDR pnmh, BOOL& bHandled );

	virtual void InitItems( void ) {}
	virtual void ItemsChanged( void ) {}
	virtual void ParseTreeItemExtra( CTreeItem *pItem, CSettingsParser &parser ) {}
	virtual void SerializeItemExtra( CTreeItem *pItem, std::vector<wchar_t> &stringBuilder ) {}
	virtual bool EditItem( CTreeItem *pItem, HWND tree, HTREEITEM hItem, std::vector<HMODULE> &modules );
	void AddItem( HTREEITEM hCommand );

	HTREEITEM GetRoot( void );
	HTREEITEM GetChild( HTREEITEM hParent );
	HTREEITEM GetNext( HTREEITEM hItem );
	CTreeItem *GetItem( HTREEITEM hItem );

	static void AppendString( std::vector<wchar_t> &stringBuilder, const wchar_t *text );

private:
	CSettingsTree &m_Tree;
	CCommandsTree &m_CommandsTree;
	const CStdCommand *m_pStdCommands;
	CSetting *m_pSetting;
	bool m_bMenu;

	const CStdCommand *FindStdCommand( const wchar_t *name );

	void EditItemInternal( CTreeItem *pItem, HTREEITEM hItem );
	int ParseTreeItem( CTreeItem *pItem, CSettingsParser &parser );
	void SerializeItem( HTREEITEM hItem, std::vector<wchar_t> &stringBuilder );
	void CreateTreeItems( CSettingsParser &parser, HTREEITEM hParent, const CSettingsParser::TreeItem *pItems, int index );
};

class CEditCustomItemDlg: public CResizeableDlg<CEditCustomItemDlg>
{
public:
	CEditCustomItemDlg( CTreeItem *pItem, std::vector<HMODULE> &modules ): m_Modules(modules) { m_pItem=pItem; }
	void SetEnableParent( HWND parent ) { m_EnableParent=parent; }
	bool GetResult( void ) { return m_bResult; }

	BEGIN_MSG_MAP( CEditCustomItemDlg )
		MESSAGE_HANDLER( WM_SIZE, OnSize )
		MESSAGE_HANDLER( WM_GETMINMAXINFO, OnGetMinMaxInfo )
	END_MSG_MAP()

	virtual BEGIN_RESIZE_MAP
	END_RESIZE_MAP

	bool Run( HWND parent, int dlgID );

protected:
	CTreeItem *m_pItem;

	// Handler prototypes:
	//  LRESULT MessageHandler(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
	//  LRESULT CommandHandler(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled);
	//  LRESULT NotifyHandler(int idCtrl, LPNMHDR pnmh, BOOL& bHandled);
	LRESULT OnSize( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled );
	LRESULT OnBrowse( WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled );
	LRESULT OnOK( WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled );
	LRESULT OnCancel( WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled );

	void InitDialog( CWindow commandCombo, const CStdCommand *pStdcommands, CWindow linkCombo, const KNOWNFOLDERID *const *pCommonLinks );
	void UpdateIcons( int iconID, int iconDID );
	CString GetComboText( WORD wNotifyCode, WORD wID );
	bool BrowseCommand( HWND parent, wchar_t *text );
	bool BrowseLink( HWND parent, wchar_t *text );
	bool BrowseIcon( wchar_t *text );

private:
	std::vector<HMODULE> m_Modules;
	HWND m_EnableParent;
	bool m_bResult;
	HICON m_hIcon;
	unsigned int m_IconKey;
	HICON m_hIconD;
	unsigned int m_IconDKey;
	CTreeItem m_StoredItem;

	void StorePlacement( void );
};

///////////////////////////////////////////////////////////////////////////////

ISettingsPanel *GetDefaultSettings( void );
ISettingsPanel *GetLanguageSettings( void );
HIMAGELIST GetSettingsImageList( HWND tree );
bool BrowseForIcon( HWND hWndParent, wchar_t *path, int &id );
const wchar_t *GetSettingsRegPath( void );

// Special GUID for the real desktop
extern const GUID FOLDERID_DesktopRoot;

enum TVersionCheck
{
	CHECK_AUTO,
	CHECK_AUTO_IE,
	CHECK_UPDATE,
};

typedef void (*tNewVersionCallback)( DWORD newVersion, CString downloadUrl, CString news );

bool CheckForNewVersion( TVersionCheck check, tNewVersionCallback callback );
