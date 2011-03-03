// Classic Shell (c) 2009-2011, Ivo Beltchev
// The sources for Classic Shell are distributed under the MIT open source license

#include "stdafx.h"
#include "resource.h"
#include "ClassicStartMenuDLL.h"
#include "Settings.h"
#include "SkinManager.h"
#include "FNVHash.h"
#include "SettingsUIHelper.h"
#include "SettingsUI.h"
#include "ResourceHelper.h"
#include "MenuContainer.h"
#include "dllmain.h"
#include <uxtheme.h>
#include <dwmapi.h>
#include <htmlhelp.h>
#define SECURITY_WIN32
#include <Security.h>

///////////////////////////////////////////////////////////////////////////////

class CSkinSettingsDlg: public CResizeableDlg<CSkinSettingsDlg>
{
public:
	CSkinSettingsDlg( bool bAllPrograms ) { m_bAllPrograms=bAllPrograms; }

	BEGIN_MSG_MAP( CSkinSettingsDlg )
		MESSAGE_HANDLER( WM_INITDIALOG, OnInitDialog )
		MESSAGE_HANDLER( WM_SIZE, OnSize )
		COMMAND_HANDLER( IDC_COMBOSKIN, CBN_SELENDOK, OnSelEndOK )
		COMMAND_HANDLER( IDC_ABOUT, BN_CLICKED, OnAbout )
		COMMAND_HANDLER( IDC_BUTTONRESET, BN_CLICKED, OnReset )
		NOTIFY_HANDLER( IDC_SKINOPTIONS, NM_CUSTOMDRAW, OnCustomDraw )
		NOTIFY_HANDLER( IDC_SKINOPTIONS, NM_CLICK, OnClick )
		NOTIFY_HANDLER( IDC_SKINOPTIONS, NM_DBLCLK, OnClick )
		NOTIFY_HANDLER( IDC_SKINOPTIONS, TVN_KEYDOWN, OnKeyDown )
		NOTIFY_HANDLER( IDC_SKINOPTIONS, TVN_GETINFOTIP, OnGetInfoTip )
	END_MSG_MAP()

	BEGIN_RESIZE_MAP
		RESIZE_CONTROL(IDC_COMBOSKIN,MOVE_SIZE_X)
		RESIZE_CONTROL(IDC_ABOUT,MOVE_MOVE_X)
		RESIZE_CONTROL(IDC_BUTTONRESET,MOVE_MOVE_X)
		RESIZE_CONTROL(IDC_STATICVER,MOVE_SIZE_X)
		RESIZE_CONTROL(IDC_SKINOPTIONS,MOVE_SIZE_X|MOVE_SIZE_Y)
		RESIZE_CONTROL(IDC_STATICALLPROGS,MOVE_SIZE_X|MOVE_MOVE_Y)
	END_RESIZE_MAP

	void SetGroup( CSetting *pGroup );

protected:
	// Handler prototypes:
	//  LRESULT MessageHandler(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
	//  LRESULT CommandHandler(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled);
	//  LRESULT NotifyHandler(int idCtrl, LPNMHDR pnmh, BOOL& bHandled);
	LRESULT OnInitDialog( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled );
	LRESULT OnSize( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled );
	LRESULT OnSelEndOK( WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled );
	LRESULT OnAbout( WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled );
	LRESULT OnReset( WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled );
	LRESULT OnCustomDraw( int idCtrl, LPNMHDR pnmh, BOOL& bHandled );
	LRESULT OnClick( int idCtrl, LPNMHDR pnmh, BOOL& bHandled );
	LRESULT OnKeyDown( int idCtrl, LPNMHDR pnmh, BOOL& bHandled );
	LRESULT OnGetInfoTip( int idCtrl, LPNMHDR pnmh, BOOL& bHandled );

private:

	struct ListOption
	{
		CString name;
		CString condition;
		unsigned int hash; // LSB is always 0
		bool bValue; // value set by the user
		bool bValue2; // default value when the condition is false
	};

	bool m_bAllPrograms;
	CWindow m_Tree;
	CSetting *m_pSetting;
	std::vector<CString> m_SkinNames;
	int m_SkinIndex;
	std::vector<CString> m_VariationNames;
	int m_VariationIndex;
	std::vector<ListOption> m_ListOptions;
	std::vector<unsigned int> m_ExtraOptions; // options that were set but are not used by the current skin
	std::vector<wchar_t> m_CurrentOptions; // a string in format XXXXXXXX|XXXXXXXX|XXXXXXXX|\0

	void InitSkinUI( void );
	void UpdateSkinSettings( void );
	void ToggleItem( HTREEITEM hItem );
};

// Subclass the tooltip to delay the tip when the mouse moves from one tree item to the next
static LRESULT CALLBACK SubclassInfoTipProc( HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData )
{
	if (uMsg==TTM_UPDATE)
	{
		int time=(int)SendMessage(hWnd,TTM_GETDELAYTIME,TTDT_RESHOW,0);
		SetTimer(hWnd,'CLSH',time,NULL);
		return 0;
	}
	if (uMsg==WM_TIMER && wParam=='CLSH')
	{
		KillTimer(hWnd,wParam);
		DefSubclassProc(hWnd,TTM_UPDATE,0,0);
		return 0;
	}
	return DefSubclassProc(hWnd,uMsg,wParam,lParam);
}

// Fills the tree with the options for the current skin. Also initializes m_ListOptions and m_CurrentOptions
void CSkinSettingsDlg::InitSkinUI( void )
{
	MenuSkin skin;
	if (!LoadMenuSkin(m_SkinNames[m_SkinIndex],skin,NULL,L"",m_bAllPrograms?0:LOADMENU_MAIN))
	{
		skin.Variations.clear();
		skin.Options.clear();
	}

	GetDlgItem(IDC_STATICVER).ShowWindow(skin.version>MAX_SKIN_VERSION?SW_SHOW:SW_HIDE);

	// store current options in m_ExtraOptions and clear the list
	int n=(int)m_ListOptions.size();
	for (int i=0;i<n;i++)
	{
		unsigned int hash=m_ListOptions[i].hash;
		bool bValue=m_ListOptions[i].bValue;
		bool bFound=false;
		for (std::vector<unsigned int>::iterator it=m_ExtraOptions.begin();it!=m_ExtraOptions.end();++it)
			if ((*it&0xFFFFFFFE)==hash)
			{
				*it=hash|(bValue?1:0);
				bFound=true;
				break;
			}
			if (!bFound)
				m_ExtraOptions.push_back(hash|(bValue?1:0));
	}

	CWindow label=GetDlgItem(IDC_STATICOPT);

	TreeView_DeleteAllItems(m_Tree);
	m_VariationNames.clear();
	m_ListOptions.clear();
	m_VariationIndex=-1;

	if (skin.Options.empty() && skin.Variations.empty())
	{
		label.ShowWindow(SW_HIDE);
		m_Tree.ShowWindow(SW_HIDE);
		m_CurrentOptions.clear();
	}
	else
	{
		// init variations
		if (!skin.Variations.empty())
		{
			const wchar_t *variaton=L"";
			if (m_pSetting[1].value.vt==VT_BSTR)
				variaton=m_pSetting[1].value.bstrVal;

			m_VariationIndex=0;
			CString vars=LoadStringEx(IDC_SKINVARIATION);
			TVINSERTSTRUCT insert={NULL,TVI_LAST,{TVIF_STATE|TVIF_TEXT|TVIF_IMAGE|TVIF_SELECTEDIMAGE|TVIF_PARAM,NULL,TVIS_EXPANDED,TVIS_EXPANDED|TVIS_OVERLAYMASK,(LPWSTR)(LPCWSTR)vars,0,SETTING_STATE_SETTING,SETTING_STATE_SETTING}};
			insert.item.lParam=(LPARAM)(m_pSetting+1);
			if (m_pSetting[1].flags&CSetting::FLAG_LOCKED_MASK)
			{
				insert.item.state|=INDEXTOOVERLAYMASK(1);
				insert.item.iImage=insert.item.iSelectedImage=SETTING_STATE_SETTING|SETTING_STATE_DISABLED;
			}
			HTREEITEM hVars=TreeView_InsertItem(m_Tree,&insert);
			int n=(int)skin.Variations.size();
			m_VariationNames.resize(n);
			for (int i=0;i<n;i++)
			{
				m_VariationNames[i]=skin.Variations[i].second;
				if (_wcsicmp(m_VariationNames[i],variaton)==0)
					m_VariationIndex=i;
				TVINSERTSTRUCT insert2={hVars,TVI_LAST,{TVIF_TEXT|TVIF_PARAM,NULL,0,0,(LPWSTR)(LPCWSTR)skin.Variations[i].second}};
				insert2.item.lParam=i;
				TreeView_InsertItem(m_Tree,&insert2);
			}
		}

		// init options
		std::vector<unsigned int> options;
		if (m_pSetting[2].value.vt==VT_BSTR)
		{
			for (const wchar_t *str=m_pSetting[2].value.bstrVal;*str;)
			{
				wchar_t token[256];
				str=GetToken(str,token,_countof(token),L"|");
				wchar_t *q;
				unsigned int hash=wcstoul(token,&q,16);
				options.push_back(hash);
			}
		}

		int n=(int)skin.Options.size();
		m_ListOptions.resize(n);
		m_CurrentOptions.resize(n*9+1);
		for (int i=0;i<n;i++)
		{
			unsigned int hash=CalcFNVHash(skin.Options[i].name)&0xFFFFFFFE;
			m_ListOptions[i].name=skin.Options[i].name;
			m_ListOptions[i].hash=hash;
			m_ListOptions[i].condition=skin.Options[i].condition;

			// get the default value
			bool bValue=skin.Options[i].value;

			// override from m_ExtraOptions
			for (std::vector<unsigned int>::const_iterator it=m_ExtraOptions.begin();it!=m_ExtraOptions.end();++it)
				if ((*it&0xFFFFFFFE)==hash)
				{
					bValue=(*it&1)==1;
					break;
				}

				// override from options
				for (std::vector<unsigned int>::const_iterator it=options.begin();it!=options.end();++it)
					if ((*it&0xFFFFFFFE)==hash)
					{
						bValue=(*it&1)==1;
						break;
					}

					m_ListOptions[i].bValue=bValue;
					Sprintf(&m_CurrentOptions[i*9],10,L"%08X|",hash|(bValue?1:0));
					m_ListOptions[i].bValue2=skin.Options[i].value2;

					TVINSERTSTRUCT insert={NULL,TVI_LAST,{TVIF_TEXT|TVIF_PARAM|TVIF_STATE,NULL,0,TVIS_OVERLAYMASK,(LPWSTR)(LPCWSTR)(LPWSTR)(LPCWSTR)skin.Options[i].label}};
					insert.item.lParam=i;
					if (m_pSetting[2].flags&CSetting::FLAG_LOCKED_MASK)
						insert.item.state|=INDEXTOOVERLAYMASK(1);
					TreeView_InsertItem(m_Tree,&insert);
		}
		UpdateSkinSettings();
		label.ShowWindow(SW_SHOW);
		m_Tree.ShowWindow(SW_SHOW);
	}
}

void CSkinSettingsDlg::UpdateSkinSettings( void )
{
	HTREEITEM hItem=TreeView_GetRoot(m_Tree);
	if (!hItem) return;
	TVITEM item={TVIF_PARAM|TVIF_IMAGE|TVIF_SELECTEDIMAGE,hItem};
	TreeView_GetItem(m_Tree,&item);
	if (item.lParam==(LPARAM)(m_pSetting+1))
	{
		// has variations
		for (HTREEITEM hVar=TreeView_GetChild(m_Tree,hItem);hVar;hVar=TreeView_GetNextSibling(m_Tree,hVar))
		{
			item.hItem=hVar;
			TreeView_GetItem(m_Tree,&item);
			int image=SETTING_STATE_RADIO;
			if ((int)item.lParam==m_VariationIndex)
				image|=SETTING_STATE_CHECKED;
			if (m_pSetting[1].flags&CSetting::FLAG_LOCKED_MASK)
				image|=SETTING_STATE_DISABLED;
			if (item.iImage!=image)
			{
				item.iImage=item.iSelectedImage=image;
				TreeView_SetItem(m_Tree,&item);
				RECT rc;
				TreeView_GetItemRect(m_Tree,hVar,&rc,FALSE);
				m_Tree.InvalidateRect(&rc);
			}
		}
		hItem=TreeView_GetNextSibling(m_Tree,hItem);
	}

	std::vector<const wchar_t*> values; // list of true values
	bool bLocked=(m_pSetting[2].flags&CSetting::FLAG_LOCKED_MASK)!=0;
	for (;hItem;hItem=TreeView_GetNextSibling(m_Tree,hItem))
	{
		item.hItem=hItem;
		TreeView_GetItem(m_Tree,&item);
		int idx=(int)item.lParam;

		bool bDisabled=(!m_ListOptions[idx].condition.IsEmpty() && !EvalCondition(m_ListOptions[idx].condition,values.empty()?NULL:&values[0],(int)values.size()));
		bool bValue=bDisabled?m_ListOptions[idx].bValue2:m_ListOptions[idx].bValue;
		if (bValue)
			values.push_back(m_ListOptions[idx].name);
		int image=SETTING_STATE_CHECKBOX;
		if (bValue)
			image|=SETTING_STATE_CHECKED;
		if (bDisabled || bLocked)
			image|=SETTING_STATE_DISABLED;
		if (item.iImage!=image)
		{
			item.iImage=item.iSelectedImage=image;
			TreeView_SetItem(m_Tree,&item);
			RECT rc;
			TreeView_GetItemRect(m_Tree,hItem,&rc,FALSE);
			m_Tree.InvalidateRect(&rc);
		}
	}
}

LRESULT CSkinSettingsDlg::OnCustomDraw( int idCtrl, LPNMHDR pnmh, BOOL& bHandled )
{
	NMTVCUSTOMDRAW *pDraw=(NMTVCUSTOMDRAW*)pnmh;
	if (pDraw->nmcd.dwDrawStage==CDDS_PREPAINT)
		return CDRF_NOTIFYITEMDRAW;
	else if (pDraw->nmcd.dwDrawStage==CDDS_ITEMPREPAINT)
	{
		TVITEM item={TVIF_IMAGE|TVIF_STATE,(HTREEITEM)pDraw->nmcd.dwItemSpec,0,TVIS_SELECTED};
		TreeView_GetItem(m_Tree,&item);
		if ((item.iImage&SETTING_STATE_DISABLED) && (!(item.state&TVIS_SELECTED) || IsAppThemed()))
			pDraw->clrText=GetSysColor(COLOR_GRAYTEXT);
	}
	return CDRF_DODEFAULT;
}

void CSkinSettingsDlg::ToggleItem( HTREEITEM hItem )
{
	if (!hItem) return;
	TVITEM item={TVIF_PARAM|TVIF_IMAGE,hItem};
	TreeView_GetItem(m_Tree,&item);
	if (item.iImage&SETTING_STATE_DISABLED)
		return;
	if (item.iImage&SETTING_STATE_RADIO)
	{
		// select variation
		if (item.iImage&SETTING_STATE_CHECKED)
			return;
		m_VariationIndex=(int)item.lParam;
		CSettingsLockWrite lock;
		m_pSetting[1].value=CComVariant(m_VariationNames[item.lParam]);
	}
	else if (item.iImage&SETTING_STATE_CHECKBOX)
	{
		wchar_t &c=m_CurrentOptions[item.lParam*9+7];
		if (c>='A') c++; // hack - 'A' is 65, which is odd. ++ makes it even
		if (item.iImage&SETTING_STATE_CHECKED)
		{
			c&=0xFFFE;
			m_ListOptions[item.lParam].bValue=false;
		}
		else
		{
			c|=1;
			m_ListOptions[item.lParam].bValue=true;
		}
		if (c>='A') c--; // unhack
		CSettingsLockWrite lock;
		m_pSetting[2].value=CComVariant(&m_CurrentOptions[0]);
	}
	UpdateSkinSettings();
}

LRESULT CSkinSettingsDlg::OnClick( int idCtrl, LPNMHDR pnmh, BOOL& bHandled )
{
	DWORD pos=GetMessagePos();
	TVHITTESTINFO test={{(short)LOWORD(pos),(short)HIWORD(pos)}};
	m_Tree.ScreenToClient(&test.pt);
	if (TreeView_HitTest(m_Tree,&test) && (test.flags&(TVHT_ONITEMICON|TVHT_ONITEMLABEL)))
		ToggleItem(test.hItem);

	return 0;
}


LRESULT CSkinSettingsDlg::OnKeyDown( int idCtrl, LPNMHDR pnmh, BOOL& bHandled )
{
	NMTVKEYDOWN *pKey=(NMTVKEYDOWN*)pnmh;
	if (pKey->wVKey==VK_SPACE)
	{
		ToggleItem(TreeView_GetSelection(m_Tree));
		return 1;
	}
	bHandled=FALSE;
	return 0;
}

LRESULT CSkinSettingsDlg::OnGetInfoTip( int idCtrl, LPNMHDR pnmh, BOOL& bHandled )
{
	NMTVGETINFOTIP *pTip=(NMTVGETINFOTIP*)pnmh;
	TVITEM item={TVIF_STATE,pTip->hItem,0,TVIS_OVERLAYMASK};
	TreeView_GetItem(m_Tree,&item);
	if (item.state&TVIS_OVERLAYMASK)
		Strcpy(pTip->pszText,pTip->cchTextMax,LoadStringEx(IDS_SETTING_LOCKED));
	return 0;
}

LRESULT CSkinSettingsDlg::OnInitDialog( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled )
{
	CResizeableDlg<CSkinSettingsDlg>::InitResize();
	EnableThemeDialogTexture(m_hWnd,ETDT_ENABLETAB);

	m_Tree=GetDlgItem(IDC_SKINOPTIONS);
	TreeView_SetImageList(m_Tree,GetSettingsImageList(m_Tree),TVSIL_NORMAL);
	if (IsAppThemed())
	{
		m_Tree.SetWindowLong(GWL_STYLE,m_Tree.GetWindowLong(GWL_STYLE)|TVS_TRACKSELECT);
		SetWindowTheme(m_Tree,L"Explorer",NULL);
	}

	m_SkinNames.clear();
	m_SkinNames.push_back(LoadStringEx(IDS_DEFAULT_SKIN));
	wchar_t find[_MAX_PATH];
	GetSkinsPath(find);
	Strcat(find,_countof(find),L"1.txt");
	if (GetFileAttributes(find)!=INVALID_FILE_ATTRIBUTES)
	{
		m_SkinNames.push_back(L"Custom");
	}

	*PathFindFileName(find)=0;
	Strcat(find,_countof(find),L"*.skin");
	WIN32_FIND_DATA data;
	HANDLE h=FindFirstFile(find,&data);
	while (h!=INVALID_HANDLE_VALUE)
	{
		if (!(data.dwFileAttributes&FILE_ATTRIBUTE_DIRECTORY))
		{
			*PathFindExtension(data.cFileName)=0;
			m_SkinNames.push_back(data.cFileName);
		}
		if (!FindNextFile(h,&data))
		{
			FindClose(h);
			break;
		}
	}

	CWindow skins=GetDlgItem(IDC_COMBOSKIN);
	for (std::vector<CString>::const_iterator it=m_SkinNames.begin();it!=m_SkinNames.end();++it)
		skins.SendMessage(CB_ADDSTRING,0,(LPARAM)(const wchar_t *)*it);

	CWindow tooltip=TreeView_GetToolTips(m_Tree);
	tooltip.SendMessage(TTM_SETDELAYTIME,TTDT_AUTOPOP,10000);
	tooltip.SendMessage(TTM_SETDELAYTIME,TTDT_INITIAL,1000);
	tooltip.SendMessage(TTM_SETDELAYTIME,TTDT_RESHOW,1000);
	SetWindowSubclass(tooltip,SubclassInfoTipProc,'CLSH',0);

	TOOLINFO tool={sizeof(tool),TTF_SUBCLASS,m_hWnd,'CLSH'};
	CString str=LoadStringEx(IDS_SETTING_LOCKED);
	tool.lpszText=(LPWSTR)(LPCWSTR)str;
	skins.GetClientRect(&tool.rect);
	skins.ClientToScreen(&tool.rect);
	ScreenToClient(&tool.rect);
	tooltip.SendMessage(TTM_ADDTOOL,0,(LPARAM)&tool);

	GetDlgItem(IDC_STATICALLPROGS).ShowWindow(m_bAllPrograms?SW_SHOW:SW_HIDE);

	m_ListOptions.clear();
	m_ExtraOptions.clear();
	m_CurrentOptions.clear();

	return TRUE;
}

LRESULT CSkinSettingsDlg::OnSize( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled )
{
	CResizeableDlg<CSkinSettingsDlg>::OnSize();

	CWindow skins=GetDlgItem(IDC_COMBOSKIN);
	CWindow tooltip=TreeView_GetToolTips(m_Tree);
	TOOLINFO tool={sizeof(tool),TTF_SUBCLASS,m_hWnd,'CLSH'};
	skins.GetClientRect(&tool.rect);
	skins.ClientToScreen(&tool.rect);
	ScreenToClient(&tool.rect);
	tooltip.SendMessage(TTM_NEWTOOLRECT,0,(LPARAM)&tool);

	return 0;
}

LRESULT CSkinSettingsDlg::OnSelEndOK( WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled )
{
	m_SkinIndex=(int)SendDlgItemMessage(IDC_COMBOSKIN,CB_GETCURSEL);
	{
		CSettingsLockWrite lock;
		m_pSetting[0].value=CComBSTR(m_SkinNames[m_SkinIndex]);
	}
	InitSkinUI();
	return 0;
}

static HRESULT CALLBACK TaskDialogCallbackProc( HWND hwnd, UINT uNotification, WPARAM wParam, LPARAM lParam, LONG_PTR dwRefData )
{
	if (uNotification==TDN_HYPERLINK_CLICKED)
	{
		ShellExecute(hwnd,L"open",(const wchar_t*)lParam,NULL,NULL,SW_SHOWNORMAL);
	}
	return S_OK;
}

LRESULT CSkinSettingsDlg::OnAbout( WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled )
{
	int idx=(int)SendDlgItemMessage(IDC_COMBOSKIN,CB_GETCURSEL,0,0);
	const wchar_t *name=m_SkinNames[idx];
	MenuSkin skin;
	wchar_t caption[256];
	Sprintf(caption,_countof(caption),LoadStringEx(IDS_SKIN_ABOUT),name);
	if (!LoadMenuSkin(name,skin,NULL,L"",0))
	{
		MessageBox(LoadStringEx(IDS_SKIN_FAIL),caption,MB_OK|MB_ICONERROR);
		return TRUE;
	}
	TASKDIALOGCONFIG task={sizeof(task),m_hWnd,NULL,TDF_ENABLE_HYPERLINKS|TDF_ALLOW_DIALOG_CANCELLATION|TDF_USE_HICON_MAIN,TDCBF_OK_BUTTON};
	task.pszWindowTitle=caption;
	task.pszContent=skin.About;
	task.hMainIcon=skin.AboutIcon?skin.AboutIcon:LoadIcon(NULL,IDI_INFORMATION);
	task.pfCallback=TaskDialogCallbackProc;
	TaskDialogIndirect(&task,NULL,NULL,NULL);
	return TRUE;
}

LRESULT CSkinSettingsDlg::OnReset( WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled )
{
	m_ListOptions.clear();
	m_CurrentOptions.clear();
	{
		CSettingsLockWrite lock;
		if (!(m_pSetting[0].flags&CSetting::FLAG_LOCKED_MASK))
			m_pSetting[0].value=m_pSetting[0].defValue;
		if (!(m_pSetting[1].flags&CSetting::FLAG_LOCKED_MASK))
			m_pSetting[1].value=m_pSetting[1].defValue;
		if (!(m_pSetting[2].flags&CSetting::FLAG_LOCKED_MASK))
			m_pSetting[2].value=m_pSetting[2].defValue;
	}
	SetGroup(m_pSetting-1);
	return TRUE;
}

void CSkinSettingsDlg::SetGroup( CSetting *pGroup )
{
	m_pSetting=pGroup+1;
	// the skin settings are never default
	{
		CSettingsLockWrite lock;
		m_pSetting[0].flags&=~CSetting::FLAG_DEFAULT;
		m_pSetting[1].flags&=~CSetting::FLAG_DEFAULT;
		m_pSetting[2].flags&=~CSetting::FLAG_DEFAULT;
	}

	const wchar_t *skin=L"";
	if (m_pSetting[0].value.vt==VT_BSTR)
		skin=m_pSetting[0].value.bstrVal;

	m_SkinIndex=-1;
	if (!*skin)
		SendDlgItemMessage(IDC_COMBOSKIN,CB_SETCURSEL,0);
	else
	{
		int n=(int)m_SkinNames.size();
		for (int i=1;i<n;i++)
		{
			if (_wcsicmp(skin,m_SkinNames[i])==0)
			{
				SendDlgItemMessage(IDC_COMBOSKIN,CB_SETCURSEL,i);
				m_SkinIndex=i;
				break;
			}
		}
	}

	if (m_SkinIndex<0)
	{
		CSettingsLockWrite lock;
		m_pSetting[0].value=CComVariant(m_SkinNames[0]);
		m_pSetting[1].value=CComVariant(L"");
		m_pSetting[2].value=CComVariant(L"");
		SendDlgItemMessage(IDC_COMBOSKIN,CB_SETCURSEL,0);
		m_SkinIndex=0;
	}
	GetDlgItem(IDC_COMBOSKIN).EnableWindow(!(m_pSetting[0].flags&CSetting::FLAG_LOCKED_MASK));

	InitSkinUI();
}

class CSkinSettingsPanel: public ISettingsPanel
{
public:
	CSkinSettingsPanel( bool bAllPrograms ) { m_bAllPrograms=bAllPrograms; }
	virtual HWND Create( HWND parent );
	virtual HWND Activate( CSetting *pGroup, const RECT &rect, bool bReset );
	virtual bool Validate( HWND parent ) { return true; }

private:
	bool m_bAllPrograms;

	static CSkinSettingsDlg s_Dialog;
	static CSkinSettingsDlg s_DialogAll;
};

CSkinSettingsDlg CSkinSettingsPanel::s_Dialog(false);
CSkinSettingsDlg CSkinSettingsPanel::s_DialogAll(true);

HWND CSkinSettingsPanel::Create( HWND parent )
{
	CSkinSettingsDlg &dlg=m_bAllPrograms?s_DialogAll:s_Dialog;
	if (!dlg.m_hWnd)
		dlg.Create(parent,LoadDialogEx(IDD_SKINSETTINGS));
	return dlg.m_hWnd;
}

HWND CSkinSettingsPanel::Activate( CSetting *pGroup, const RECT &rect, bool bReset )
{
	CSkinSettingsDlg &dlg=m_bAllPrograms?s_DialogAll:s_Dialog;
	dlg.SetGroup(pGroup);
	dlg.SetWindowPos(HWND_TOP,&rect,SWP_SHOWWINDOW);
	return dlg.m_hWnd;
}

static CSkinSettingsPanel g_SkinSettingsPanel(false);
static CSkinSettingsPanel g_SkinSettingsPanelAll(true);

///////////////////////////////////////////////////////////////////////////////

static CStdCommand g_StdCommands[]={
	{L"SEPARATOR",L"SEPARATOR",IDS_SEPARATOR_TIP},
	{L"COLUMN_BREAK",L"COLUMN_BREAK",IDS_BREAK_TIP},
	{L"COLUMN_PADDING",L"COLUMN_PADDING",IDS_PADDING_TIP},
	{L"programs",L"Programs",IDS_PROGRAMS_TIP,L"ProgramsMenu",L"$Menu.Programs",L"",L"shell32.dll,326",&FOLDERID_Programs,NULL,StdMenuItem::MENU_TRACK},
	{L"favorites",L"Favorites",IDS_FAVORITES_TIP,L"FavoritesItem",L"$Menu.Favorites",L"",L"shell32.dll,322",&FOLDERID_Favorites},
	{L"documents",L"Documents",IDS_DOCUMENTS_TIP,L"DocumentsItem",L"$Menu.Documents",L"",L"shell32.dll,327",&FOLDERID_Recent,NULL,StdMenuItem::MENU_ITEMS_FIRST},
	{L"settings",L"Settings",IDS_SETTINGS_MENU_TIP,L"SettingsMenu",L"$Menu.Settings",L"",L"shell32.dll,330"},
	{L"search",L"Search Menu",IDS_SEARCH_TIP,L"SearchMenu",L"$Menu.Search",L"",L"shell32.dll,323"},
	{L"search_box",L"Search Box",IDS_SEARCH_BOX_TIP,L"SearchBoxItem",L"$Menu.SearchBox",NULL,L"none"},
	{L"help",L"Help",IDS_HELP_TIP,L"HelpItem",L"$Menu.Help",NULL,L"shell32.dll,324"},
	{L"run",L"Run",IDS_RUN_TIP,L"RunItem",L"$Menu.Run",NULL,L"shell32.dll,328"},
	{L"logoff",L"Log Off",IDS_LOGOFF_TIP,L"LogOffItem",L"$Menu.Logoff",NULL,L"shell32.dll,325"},
	{L"undock",L"Undock",IDS_UNDOCK_TIP,L"UndockItem",L"$Menu.Undock",NULL,L"shell32.dll,331"},
	{L"disconnect",L"Disconnect",IDS_DISCONNECT_TIP,L"DisconnectItem",L"$Menu.Disconnect",NULL,L"shell32.dll,329"},
	{L"shutdown_box",L"Shutdown Box",IDS_SHUTDOWNBOX_TIP,L"ShutdownBoxItem",L"$Menu.ShutdownBox",NULL,L"shell32.dll,329"},
	{L"user_files",L"User Files",IDS_USERFILES_TIP,L"UserFilesItem",NULL,L"$Menu.UserFilesTip",L"",&FOLDERID_UsersFiles},
	{L"user_documents",L"User Documents",IDS_USERDOCS_TIP,L"UserDocumentsItem",NULL,L"$Menu.UserDocumentsTip",L"",&FOLDERID_Documents},
	{L"user_pictures",L"User Pictures",IDS_USERPICS_TIP,L"UserPicturesItem",NULL,L"$Menu.UserPicturesTip",L"",&FOLDERID_Pictures},
	{L"control_panel",L"Control Panel",IDS_CONTROLPANEL_TIP,L"ControlPanelItem",L"$Menu.ControlPanel",NULL,L"shell32.dll,137",&FOLDERID_ControlPanelFolder},
	{L"control_panel_categories",L"Control Panel Categories",IDS_CONTROLPANEL2_TIP,L"ControlPanelItem",L"$Menu.ControlPanel",NULL,L"shell32.dll,137",&FOLDERID_ControlPanelFolder},
	{L"windows_security",L"Windows Security",IDS_SECURITY_TIP,L"SecurityItem",L"$Menu.Security",NULL,L"shell32.dll,48"},
	{L"network_connections",L"Network Connections",IDS_NETWORK_TIP,L"NetworkItem",L"$Menu.Network",L"$Menu.NetworkTip",L"shell32.dll,257",&FOLDERID_ConnectionsFolder},
	{L"printers",L"Printers",IDS_PRINTERS_TIP,L"PrintersItem",L"$Menu.Printers",L"$Menu.PrintersTip",L"shell32.dll,138",&FOLDERID_PrintersFolder},
	{L"taskbar_settings",L"Taskbar Settings",IDS_TASKBAR_TIP,L"TaskbarSettingsItem",L"$Menu.Taskbar",L"$Menu.TaskbarTip",L"shell32.dll,40"},
	{L"programs_features",L"Programs and Features",IDS_FEATURES_TIP,L"ProgramsFeaturesItem",L"$Menu.Features",L"$Menu.FeaturesTip",L"shell32.dll,271"},
	{L"menu_settings",L"Menu Settings",IDS_MENU_TIP,L"MenuSettingsItem",L"$Menu.ClassicSettings",L"$Menu.SettingsTip",L",1"},
	{L"search_files",L"Search Files",IDS_SEARCHFI_TIP,L"SearchFilesItem",L"$Menu.SearchFiles",NULL,L"shell32.dll,134"},
	{L"search_printer",L"Search Printer",IDS_SEARCHPR_TIP,L"SearchPrinterItem",L"$Menu.SearchPrinter",NULL,L"shell32.dll,1006"},
	{L"search_computers",L"Search Computers",IDS_SEARCHCO_TIP,L"SearchComputersItem",L"$Menu.SearchComputers",NULL,L"shell32.dll,135"},
	{L"search_people",L"Search People",IDS_SEARCHPE_TIP,L"SearchPeopleItem",L"$Menu.SearchPeople",NULL,L"shell32.dll,269"},
	{L"sleep",L"Sleep",IDS_SLEEP_TIP,L"SleepItem",L"$Menu.Sleep",NULL,L""},
	{L"hibernate",L"Hibernate",IDS_HIBERNATE_TIP,L"HibernateItem",L"$Menu.Hibernate",NULL,L""},
	{L"restart",L"Restart",IDS_RESTART_TIP,L"RestartItem",L"$Menu.Restart",NULL,L""},
	{L"shutdown",L"Shutdown",IDS_SHUTDOWN_TIP,L"ShutdownItem",L"$Menu.Shutdown",NULL,L""},
	{L"switch_user",L"Switch User",IDS_SWITCH_TIP,L"SwitchUserItem",L"$Menu.SwitchUser",NULL,L""},
	{L"recent_items",L"Recent Items",IDS_RECENT_TIP,L"RecentItems",NULL,NULL,L""},
	{L"",L"Custom Command",IDS_CUSTOM_TIP,L"CustomItem",NULL,NULL,L""},
	{NULL},
};

const wchar_t *g_DefaultStartMenu=
L"Items=COLUMN_PADDING, ProgramsMenu, COLUMN_BREAK, FavoritesItem, DocumentsItem, SettingsMenu, SearchMenu, HelpItem, RunItem, SEPARATOR, LogOffItem, UndockItem, DisconnectItem, ShutdownBoxItem, SearchBoxItem\n"
L"ProgramsMenu.Command=programs\n"
L"ProgramsMenu.Label=$Menu.Programs\n"
L"ProgramsMenu.Icon=shell32.dll,326\n"
L"FavoritesItem.Command=favorites\n"
L"FavoritesItem.Label=$Menu.Favorites\n"
L"FavoritesItem.Icon=shell32.dll,322\n"
L"DocumentsItem.Command=documents\n"
L"DocumentsItem.Items=UserFilesItem, UserDocumentsItem, UserPicturesItem\n"
L"DocumentsItem.Label=$Menu.Documents\n"
L"DocumentsItem.Icon=shell32.dll,327\n"
L"DocumentsItem.Settings=ITEMS_FIRST\n"
L"SettingsMenu.Command=settings\n"
L"SettingsMenu.Items=ControlPanelItem, SEPARATOR, SecurityItem, NetworkItem, PrintersItem, TaskbarSettingsItem, ProgramsFeaturesItem, SEPARATOR, MenuSettingsItem\n"
L"SettingsMenu.Label=$Menu.Settings\n"
L"SettingsMenu.Icon=shell32.dll,330\n"
L"SearchMenu.Command=search\n"
L"SearchMenu.Items=SearchFilesItem, SearchPrinterItem, SearchComputersItem, SearchPeopleItem\n"
L"SearchMenu.Label=$Menu.Search\n"
L"SearchMenu.Icon=shell32.dll,323\n"
L"HelpItem.Command=help\n"
L"HelpItem.Label=$Menu.Help\n"
L"HelpItem.Icon=shell32.dll,324\n"
L"RunItem.Command=run\n"
L"RunItem.Label=$Menu.Run\n"
L"RunItem.Icon=shell32.dll,328\n"
L"LogOffItem.Command=logoff\n"
L"LogOffItem.Label=$Menu.Logoff\n"
L"LogOffItem.Icon=shell32.dll,325\n"
L"UndockItem.Command=undock\n"
L"UndockItem.Label=$Menu.Undock\n"
L"UndockItem.Icon=shell32.dll,331\n"
L"DisconnectItem.Command=disconnect\n"
L"DisconnectItem.Label=$Menu.Disconnect\n"
L"DisconnectItem.Icon=shell32.dll,329\n"
L"ShutdownBoxItem.Command=shutdown_box\n"
L"ShutdownBoxItem.Label=$Menu.ShutdownBox\n"
L"ShutdownBoxItem.Icon=shell32.dll,329\n"
L"ShutdownBoxItem.Items=SwitchUserItem, SleepItem, HibernateItem, RestartItem, ShutdownItem\n"
L"SearchBoxItem.Command=search_box\n"
L"SearchBoxItem.Label=$Menu.SearchBox\n"
L"SearchBoxItem.Icon=none\n"
L"SearchBoxItem.Settings=TRACK_RECENT, OPEN_UP\n"
L"UserFilesItem.Command=user_files\n"
L"UserFilesItem.Tip=$Menu.UserFilesTip\n"
L"UserDocumentsItem.Command=user_documents\n"
L"UserDocumentsItem.Tip=$Menu.UserDocumentsTip\n"
L"UserPicturesItem.Command=user_pictures\n"
L"UserPicturesItem.Tip=$Menu.UserPicturesTip\n"
L"ControlPanelItem.Command=control_panel\n"
L"ControlPanelItem.Icon=shell32.dll,137\n"
L"ControlPanelItem.Label=$Menu.ControlPanel\n"
L"SecurityItem.Command=windows_security\n"
L"SecurityItem.Icon=shell32.dll,48\n"
L"SecurityItem.Label=$Menu.Security\n"
L"NetworkItem.Command=network_connections\n"
L"NetworkItem.Icon=shell32.dll,257\n"
L"NetworkItem.Label=$Menu.Network\n"
L"NetworkItem.Tip=$Menu.NetworkTip\n"
L"PrintersItem.Command=printers\n"
L"PrintersItem.Icon=shell32.dll,138\n"
L"PrintersItem.Label=$Menu.Printers\n"
L"PrintersItem.Tip=$Menu.PrintersTip\n"
L"TaskbarSettingsItem.Command=taskbar_settings\n"
L"TaskbarSettingsItem.Label=$Menu.Taskbar\n"
L"TaskbarSettingsItem.Icon=shell32.dll,40\n"
L"TaskbarSettingsItem.Tip=$Menu.TaskbarTip\n"
L"ProgramsFeaturesItem.Command=programs_features\n"
L"ProgramsFeaturesItem.Label=$Menu.Features\n"
L"ProgramsFeaturesItem.Icon=shell32.dll,271\n"
L"ProgramsFeaturesItem.Tip=$Menu.FeaturesTip\n"
L"MenuSettingsItem.Command=menu_settings\n"
L"MenuSettingsItem.Label=$Menu.ClassicSettings\n"
L"MenuSettingsItem.Icon=,1\n"
L"MenuSettingsItem.Tip=$Menu.SettingsTip\n"
L"SearchFilesItem.Command=search_files\n"
L"SearchFilesItem.Label=$Menu.SearchFiles\n"
L"SearchFilesItem.Icon=shell32.dll,134\n"
L"SearchPrinterItem.Command=search_printer\n"
L"SearchPrinterItem.Label=$Menu.SearchPrinter\n"
L"SearchPrinterItem.Icon=shell32.dll,1006\n"
L"SearchComputersItem.Command=search_computers\n"
L"SearchComputersItem.Label=$Menu.SearchComputers\n"
L"SearchComputersItem.Icon=shell32.dll,135\n"
L"SearchPeopleItem.Command=search_people\n"
L"SearchPeopleItem.Label=$Menu.SearchPeople\n"
L"SearchPeopleItem.Icon=shell32.dll,269\n"
L"SwitchUserItem.Command=switch_user\n"
L"SwitchUserItem.Label=$Menu.SwitchUser\n"
L"SwitchUserItem.Icon=none\n"
L"SleepItem.Command=sleep\n"
L"SleepItem.Label=$Menu.Sleep\n"
L"SleepItem.Icon=none\n"
L"HibernateItem.Command=hibernate\n"
L"HibernateItem.Label=$Menu.Hibernate\n"
L"HibernateItem.Icon=none\n"
L"RestartItem.Command=restart\n"
L"RestartItem.Label=$Menu.Restart\n"
L"RestartItem.Icon=none\n"
L"ShutdownItem.Command=shutdown\n"
L"ShutdownItem.Label=$Menu.Shutdown\n"
L"ShutdownItem.Icon=none\n"
;

// Define some Windows 7 GUIDs manually, so we don't need the Windows 7 SDK to compile Classic Shell
static const GUID FOLDERID_HomeGroup2={0x52528a6b, 0xb9e3, 0x4add, {0xb6, 0x0d, 0x58, 0x8c, 0x2d, 0xba, 0x84, 0x2d}};
static const GUID FOLDERID_Libraries2={0x1B3EA5DC, 0xB587, 0x4786, {0xb4, 0xef, 0xbd, 0x1d, 0xc3, 0x32, 0xae, 0xae}};
static const GUID FOLDERID_DocumentsLibrary2={0x7B0DB17D, 0x9CD2, 0x4A93, {0x97, 0x33, 0x46, 0xcc, 0x89, 0x02, 0x2e, 0x7c}};
static const GUID FOLDERID_MusicLibrary2={0x2112AB0A, 0xC86A, 0x4FFE, {0xa3, 0x68, 0x0d, 0xe9, 0x6e, 0x47, 0x01, 0x2e}};
static const GUID FOLDERID_PicturesLibrary2={0xA990AE9F, 0xA03B, 0x4E80, {0x94, 0xbc, 0x99, 0x12, 0xd7, 0x50, 0x41, 0x04}};
static const GUID FOLDERID_VideosLibrary2={0x491E922F, 0x5643, 0x4AF4, {0xa7, 0xeb, 0x4e, 0x7a, 0x13, 0x8d, 0x81, 0x74}};

static const KNOWNFOLDERID *g_CommonLinks[]=
{
	&FOLDERID_CommonAdminTools,
	&FOLDERID_ComputerFolder,
	&FOLDERID_DesktopRoot,
	&FOLDERID_Desktop,
	&FOLDERID_Documents,
	&FOLDERID_Downloads,
	&FOLDERID_Fonts,
	&FOLDERID_Games,
	&FOLDERID_Links,
	&FOLDERID_Music,
	&FOLDERID_Pictures,
	&FOLDERID_RecycleBinFolder,
	&FOLDERID_Videos,
	&FOLDERID_Profile,
	&FOLDERID_HomeGroup2,
	&FOLDERID_Libraries2,
	&FOLDERID_DocumentsLibrary2,
	&FOLDERID_MusicLibrary2,
	&FOLDERID_PicturesLibrary2,
	&FOLDERID_VideosLibrary2,
	&FOLDERID_NetworkFolder,
	NULL,
};

///////////////////////////////////////////////////////////////////////////////

class CEditMenuDlg: public CEditCustomItemDlg
{
public:
	CEditMenuDlg( CTreeItem *pItem, std::vector<HMODULE> &modules ): CEditCustomItemDlg(pItem,modules) {}

	BEGIN_MSG_MAP( CEditMenuDlg )
		MESSAGE_HANDLER( WM_INITDIALOG, OnInitDialog )
		COMMAND_ID_HANDLER( IDOK, OnOK )
		COMMAND_ID_HANDLER( IDCANCEL, OnCancel )
		COMMAND_HANDLER( IDC_COMBOCOMMAND, CBN_KILLFOCUS, OnCommandChanged )
		COMMAND_HANDLER( IDC_COMBOCOMMAND, CBN_SELENDOK, OnCommandChanged )
		COMMAND_HANDLER( IDC_BUTTONCOMMAND, BN_CLICKED, OnBrowseCommand )
		COMMAND_HANDLER( IDC_BUTTONLINK, BN_CLICKED, OnBrowseLink )
		COMMAND_HANDLER( IDC_BUTTONICON, BN_CLICKED, OnBrowseIcon )
		COMMAND_HANDLER( IDC_COMBOLINK, CBN_KILLFOCUS, OnLinkChanged )
		COMMAND_HANDLER( IDC_COMBOLINK, CBN_SELENDOK, OnLinkChanged )
		COMMAND_HANDLER( IDC_EDITICON, EN_KILLFOCUS, OnIconChanged )
		COMMAND_HANDLER( IDC_CHECKTRACK, BN_CLICKED, OnCheckTrack )
		COMMAND_HANDLER( IDC_CHECKNOTRACK, BN_CLICKED, OnCheckTrack )
		COMMAND_HANDLER( IDC_CHECKMULTICOLUMN, BN_CLICKED, OnCheckMulti )
		COMMAND_HANDLER( IDC_BUTTONRESET, BN_CLICKED, OnReset )
		CHAIN_MSG_MAP( CEditCustomItemDlg )
	END_MSG_MAP()

	virtual BEGIN_RESIZE_MAP
	RESIZE_CONTROL(IDC_COMBOCOMMAND,MOVE_SIZE_X)
	RESIZE_CONTROL(IDC_COMBOLINK,MOVE_SIZE_X)
	RESIZE_CONTROL(IDC_BUTTONCOMMAND,MOVE_MOVE_X)
	RESIZE_CONTROL(IDC_BUTTONLINK,MOVE_MOVE_X)
	RESIZE_CONTROL(IDC_EDITLABEL,MOVE_SIZE_X)
	RESIZE_CONTROL(IDC_EDITTIP,MOVE_SIZE_X)
	RESIZE_CONTROL(IDC_EDITICON,MOVE_SIZE_X)
	RESIZE_CONTROL(IDC_BUTTONICON,MOVE_MOVE_X)
	RESIZE_CONTROL(IDOK,MOVE_MOVE_X)
	RESIZE_CONTROL(IDCANCEL,MOVE_MOVE_X)
	END_RESIZE_MAP

protected:
	// Handler prototypes:
	//  LRESULT MessageHandler(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
	//  LRESULT CommandHandler(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled);
	//  LRESULT NotifyHandler(int idCtrl, LPNMHDR pnmh, BOOL& bHandled);
	LRESULT OnInitDialog( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled );
	LRESULT OnOK( WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled );
	LRESULT OnCommandChanged( WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled );
	LRESULT OnLinkChanged( WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled );
	LRESULT OnIconChanged( WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled );
	LRESULT OnCheckTrack( WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled );
	LRESULT OnCheckMulti( WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled );
	LRESULT OnBrowseCommand( WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled );
	LRESULT OnBrowseLink( WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled );
	LRESULT OnBrowseIcon( WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled );
	LRESULT OnReset( WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled );
};

LRESULT CEditMenuDlg::OnInitDialog( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled )
{
	CWindow commands=GetDlgItem(IDC_COMBOCOMMAND);
	CWindow links=GetDlgItem(IDC_COMBOLINK);
	InitDialog(commands,g_StdCommands,links,g_CommonLinks);
	SetDlgItemText(IDC_EDITLABEL,m_pItem->label);
	SetDlgItemText(IDC_EDITTIP,m_pItem->tip);
	SetDlgItemText(IDC_EDITICON,m_pItem->icon);

	BOOL bEnable=!(m_pItem->pStdCommand && m_pItem->pStdCommand->knownFolder);
	links.EnableWindow(bEnable);
	GetDlgItem(IDC_BUTTONLINK).EnableWindow(bEnable);
	GetDlgItem(IDC_BUTTONRESET).EnableWindow(m_pItem->pStdCommand && *m_pItem->pStdCommand->name);

	if (m_pItem->settings&StdMenuItem::MENU_TRACK)
		m_pItem->settings&=~StdMenuItem::MENU_NOTRACK;
	CheckDlgButton(IDC_CHECKSORTZA,(m_pItem->settings&StdMenuItem::MENU_SORTZA)?BST_CHECKED:BST_UNCHECKED);
	CheckDlgButton(IDC_CHECKSORTZAREC,(m_pItem->settings&StdMenuItem::MENU_SORTZA_REC)?BST_CHECKED:BST_UNCHECKED);
	CheckDlgButton(IDC_CHECKSORTONCE,(m_pItem->settings&StdMenuItem::MENU_SORTONCE)?BST_CHECKED:BST_UNCHECKED);
	CheckDlgButton(IDC_CHECKOPENUP,(m_pItem->settings&StdMenuItem::MENU_OPENUP)?BST_CHECKED:BST_UNCHECKED);
	CheckDlgButton(IDC_CHECKOPENUPREC,(m_pItem->settings&StdMenuItem::MENU_OPENUP_REC)?BST_CHECKED:BST_UNCHECKED);
	CheckDlgButton(IDC_CHECKNOEXPAND,(m_pItem->settings&StdMenuItem::MENU_NOEXPAND)?BST_CHECKED:BST_UNCHECKED);
	CheckDlgButton(IDC_CHECKNOTRACK,(m_pItem->settings&StdMenuItem::MENU_NOTRACK)?BST_CHECKED:BST_UNCHECKED);
	CheckDlgButton(IDC_CHECKITEMSFIRST,(m_pItem->settings&StdMenuItem::MENU_ITEMS_FIRST)?BST_CHECKED:BST_UNCHECKED);
	CheckDlgButton(IDC_CHECKINLINE,(m_pItem->settings&StdMenuItem::MENU_INLINE)?BST_CHECKED:BST_UNCHECKED);
	CheckDlgButton(IDC_CHECKNOEXT,(m_pItem->settings&StdMenuItem::MENU_NOEXTENSIONS)?BST_CHECKED:BST_UNCHECKED);
	if (m_pItem->pStdCommand && wcscmp(m_pItem->pStdCommand->name,L"programs")==0)
	{
		CheckDlgButton(IDC_CHECKMULTICOLUMN,BST_CHECKED);
		GetDlgItem(IDC_CHECKMULTICOLUMN).EnableWindow(FALSE);
		CheckDlgButton(IDC_CHECKTRACK,(m_pItem->settings&StdMenuItem::MENU_NOTRACK)?BST_UNCHECKED:BST_CHECKED);
	}
	else
	{
		CheckDlgButton(IDC_CHECKMULTICOLUMN,(m_pItem->settings&StdMenuItem::MENU_MULTICOLUMN)?BST_CHECKED:BST_UNCHECKED);
		CheckDlgButton(IDC_CHECKTRACK,(m_pItem->settings&StdMenuItem::MENU_TRACK)?BST_CHECKED:BST_UNCHECKED);
	}

	UpdateIcons(IDC_ICONN,0);
	SendDlgItemMessage(IDC_EDITLABEL,EM_SETCUEBANNER,TRUE,(LPARAM)(const wchar_t*)LoadStringEx(IDS_NO_TEXT));

	CWindow tooltip=CreateWindowEx(WS_EX_TOPMOST|WS_EX_TOOLWINDOW|WS_EX_TRANSPARENT,TOOLTIPS_CLASS,NULL,WS_POPUP|TTS_NOPREFIX|TTS_ALWAYSTIP,0,0,0,0,m_hWnd,NULL,g_Instance,NULL);
	tooltip.SendMessage(TTM_SETMAXTIPWIDTH,0,GetSystemMetrics(SM_CXSCREEN)/2);

	TOOLINFO tool={sizeof(tool),TTF_SUBCLASS|TTF_IDISHWND,m_hWnd,(UINT_PTR)(HWND)commands};
	CString str=LoadStringEx(IDS_COMMAND_TIP);
	tool.lpszText=(LPWSTR)(LPCWSTR)str;
	tooltip.SendMessage(TTM_ADDTOOL,0,(LPARAM)&tool);
	tool.uId=(UINT_PTR)(HWND)commands.GetWindow(GW_CHILD);
	tooltip.SendMessage(TTM_ADDTOOL,0,(LPARAM)&tool);

	str=LoadStringEx(IDS_LINK_TIP);
	tool.lpszText=(LPWSTR)(LPCWSTR)str;
	tool.uId=(UINT_PTR)(HWND)links;
	tooltip.SendMessage(TTM_ADDTOOL,0,(LPARAM)&tool);
	tool.uId=(UINT_PTR)(HWND)links.GetWindow(GW_CHILD);
	tooltip.SendMessage(TTM_ADDTOOL,0,(LPARAM)&tool);

	str=LoadStringEx(IDS_TEXT_TIP);
	tool.lpszText=(LPWSTR)(LPCWSTR)str;
	tool.uId=(UINT_PTR)(HWND)GetDlgItem(IDC_EDITLABEL);
	tooltip.SendMessage(TTM_ADDTOOL,0,(LPARAM)&tool);

	str=LoadStringEx(IDS_TIP_TIP);
	tool.lpszText=(LPWSTR)(LPCWSTR)str;
	tool.uId=(UINT_PTR)(HWND)GetDlgItem(IDC_EDITTIP);
	tooltip.SendMessage(TTM_ADDTOOL,0,(LPARAM)&tool);

	str=LoadStringEx(IDS_ICON_TIP);
	tool.lpszText=(LPWSTR)(LPCWSTR)str;
	tool.uId=(UINT_PTR)(HWND)GetDlgItem(IDC_EDITICON);
	tooltip.SendMessage(TTM_ADDTOOL,0,(LPARAM)&tool);

	str=LoadStringEx(IDS_SORTZA_TIP);
	tool.lpszText=(LPWSTR)(LPCWSTR)str;
	tool.uId=(UINT_PTR)(HWND)GetDlgItem(IDC_CHECKSORTZA);
	tooltip.SendMessage(TTM_ADDTOOL,0,(LPARAM)&tool);

	str=LoadStringEx(IDS_SORTZAREC_TIP);
	tool.lpszText=(LPWSTR)(LPCWSTR)str;
	tool.uId=(UINT_PTR)(HWND)GetDlgItem(IDC_CHECKSORTZAREC);
	tooltip.SendMessage(TTM_ADDTOOL,0,(LPARAM)&tool);

	str=LoadStringEx(IDS_SORTONCE_TIP);
	tool.lpszText=(LPWSTR)(LPCWSTR)str;
	tool.uId=(UINT_PTR)(HWND)GetDlgItem(IDC_CHECKSORTONCE);
	tooltip.SendMessage(TTM_ADDTOOL,0,(LPARAM)&tool);

	str=LoadStringEx(IDS_OPENUP_TIP);
	tool.lpszText=(LPWSTR)(LPCWSTR)str;
	tool.uId=(UINT_PTR)(HWND)GetDlgItem(IDC_CHECKOPENUP);
	tooltip.SendMessage(TTM_ADDTOOL,0,(LPARAM)&tool);

	str=LoadStringEx(IDS_OPENUPREC_TIP);
	tool.lpszText=(LPWSTR)(LPCWSTR)str;
	tool.uId=(UINT_PTR)(HWND)GetDlgItem(IDC_CHECKOPENUPREC);
	tooltip.SendMessage(TTM_ADDTOOL,0,(LPARAM)&tool);

	str=LoadStringEx(IDS_NOEXPAND_TIP);
	tool.lpszText=(LPWSTR)(LPCWSTR)str;
	tool.uId=(UINT_PTR)(HWND)GetDlgItem(IDC_CHECKNOEXPAND);
	tooltip.SendMessage(TTM_ADDTOOL,0,(LPARAM)&tool);

	str=LoadStringEx(IDS_MULTICOLUMN_TIP);
	tool.lpszText=(LPWSTR)(LPCWSTR)str;
	tool.uId=(UINT_PTR)(HWND)GetDlgItem(IDC_CHECKMULTICOLUMN);
	tooltip.SendMessage(TTM_ADDTOOL,0,(LPARAM)&tool);

	str=LoadStringEx(IDS_TRACK_TIP);
	tool.lpszText=(LPWSTR)(LPCWSTR)str;
	tool.uId=(UINT_PTR)(HWND)GetDlgItem(IDC_CHECKTRACK);
	tooltip.SendMessage(TTM_ADDTOOL,0,(LPARAM)&tool);

	str=LoadStringEx(IDS_NOTRACK_TIP);
	tool.lpszText=(LPWSTR)(LPCWSTR)str;
	tool.uId=(UINT_PTR)(HWND)GetDlgItem(IDC_CHECKNOTRACK);
	tooltip.SendMessage(TTM_ADDTOOL,0,(LPARAM)&tool);

	str=LoadStringEx(IDS_ITEMSFIRST_TIP);
	tool.lpszText=(LPWSTR)(LPCWSTR)str;
	tool.uId=(UINT_PTR)(HWND)GetDlgItem(IDC_CHECKITEMSFIRST);
	tooltip.SendMessage(TTM_ADDTOOL,0,(LPARAM)&tool);

	str=LoadStringEx(IDS_INLINE_TIP);
	tool.lpszText=(LPWSTR)(LPCWSTR)str;
	tool.uId=(UINT_PTR)(HWND)GetDlgItem(IDC_CHECKINLINE);
	tooltip.SendMessage(TTM_ADDTOOL,0,(LPARAM)&tool);

	str=LoadStringEx(IDS_NOEXTENSIONS_TIP);
	tool.lpszText=(LPWSTR)(LPCWSTR)str;
	tool.uId=(UINT_PTR)(HWND)GetDlgItem(IDC_CHECKNOEXT);
	tooltip.SendMessage(TTM_ADDTOOL,0,(LPARAM)&tool);

	str=LoadStringEx(IDS_RESTORE_TIP);
	tool.lpszText=(LPWSTR)(LPCWSTR)str;
	tool.uId=(UINT_PTR)(HWND)GetDlgItem(IDC_BUTTONRESET);
	tooltip.SendMessage(TTM_ADDTOOL,0,(LPARAM)&tool);
	return TRUE;
}

LRESULT CEditMenuDlg::OnOK( WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled )
{
	GetDlgItemText(IDC_EDITLABEL,m_pItem->label);
	m_pItem->label.TrimLeft();
	m_pItem->label.TrimRight();
	GetDlgItemText(IDC_EDITTIP,m_pItem->tip);
	m_pItem->tip.TrimLeft();
	m_pItem->tip.TrimRight();

	m_pItem->settings=0;
	bool bPrograms=(m_pItem->pStdCommand && wcscmp(m_pItem->pStdCommand->name,L"programs")==0);
	if (IsDlgButtonChecked(IDC_CHECKSORTZA)==BST_CHECKED) m_pItem->settings|=StdMenuItem::MENU_SORTZA;
	if (IsDlgButtonChecked(IDC_CHECKSORTZAREC)==BST_CHECKED) m_pItem->settings|=StdMenuItem::MENU_SORTZA_REC;
	if (IsDlgButtonChecked(IDC_CHECKSORTONCE)==BST_CHECKED) m_pItem->settings|=StdMenuItem::MENU_SORTONCE;
	if (IsDlgButtonChecked(IDC_CHECKOPENUP)==BST_CHECKED) m_pItem->settings|=StdMenuItem::MENU_OPENUP;
	if (IsDlgButtonChecked(IDC_CHECKOPENUPREC)==BST_CHECKED) m_pItem->settings|=StdMenuItem::MENU_OPENUP_REC;
	if (IsDlgButtonChecked(IDC_CHECKNOEXPAND)==BST_CHECKED) m_pItem->settings|=StdMenuItem::MENU_NOEXPAND;
	if (IsDlgButtonChecked(IDC_CHECKTRACK)==BST_CHECKED) m_pItem->settings|=StdMenuItem::MENU_TRACK;
	if (IsDlgButtonChecked(IDC_CHECKNOTRACK)==BST_CHECKED) m_pItem->settings|=StdMenuItem::MENU_NOTRACK;
	if (IsDlgButtonChecked(IDC_CHECKITEMSFIRST)==BST_CHECKED) m_pItem->settings|=StdMenuItem::MENU_ITEMS_FIRST;
	if (IsDlgButtonChecked(IDC_CHECKINLINE)==BST_CHECKED) m_pItem->settings|=StdMenuItem::MENU_INLINE;
	if (IsDlgButtonChecked(IDC_CHECKNOEXT)==BST_CHECKED) m_pItem->settings|=StdMenuItem::MENU_NOEXTENSIONS;
	if (bPrograms)
	{
		// special handling of the Programs menu
		// it is always MULTICOLUMN
		// it is always tracking, unless NOTRACK is set
		m_pItem->settings&=~StdMenuItem::MENU_TRACK;
	}
	else
	{
		if (IsDlgButtonChecked(IDC_CHECKMULTICOLUMN)==BST_CHECKED)
			m_pItem->settings|=StdMenuItem::MENU_MULTICOLUMN;
	}

	return CEditCustomItemDlg::OnOK(wNotifyCode,wID,hWndCtl,bHandled);
}

LRESULT CEditMenuDlg::OnCommandChanged( WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled )
{
	CString text=GetComboText(wNotifyCode,wID);
	if (text==m_pItem->command) return 0;
	m_pItem->SetCommand(text,g_StdCommands);
	BOOL bEnable=!(m_pItem->pStdCommand && m_pItem->pStdCommand->knownFolder);
	GetDlgItem(IDC_COMBOLINK).EnableWindow(bEnable);
	GetDlgItem(IDC_BUTTONLINK).EnableWindow(bEnable);
	GetDlgItem(IDC_BUTTONRESET).EnableWindow(m_pItem->pStdCommand && *m_pItem->pStdCommand->name);
	if (m_pItem->pStdCommand && wcscmp(m_pItem->pStdCommand->name,L"programs")==0)
	{
		CheckDlgButton(IDC_CHECKMULTICOLUMN,BST_CHECKED);
		GetDlgItem(IDC_CHECKMULTICOLUMN).EnableWindow(FALSE);
	}
	else
	{
		CheckDlgButton(IDC_CHECKMULTICOLUMN,(m_pItem->settings&StdMenuItem::MENU_MULTICOLUMN)?BST_CHECKED:BST_UNCHECKED);
		GetDlgItem(IDC_CHECKMULTICOLUMN).EnableWindow(TRUE);
	}
	UpdateIcons(IDC_ICONN,0);
	return 0;
}

LRESULT CEditMenuDlg::OnLinkChanged( WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled )
{
	CString text=GetComboText(wNotifyCode,wID);
	if (text==m_pItem->link) return 0;
	m_pItem->link=text;
	UpdateIcons(IDC_ICONN,0);
	return 0;
}

LRESULT CEditMenuDlg::OnIconChanged( WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled )
{
	CString text;
	GetDlgItemText(IDC_EDITICON,text);
	text.TrimLeft();
	text.TrimRight();
	if (text==m_pItem->icon) return 0;
	m_pItem->icon=text;
	UpdateIcons(IDC_ICONN,0);
	return 0;
}

LRESULT CEditMenuDlg::OnCheckTrack( WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled )
{
	if (IsDlgButtonChecked(wID)==BST_CHECKED)
		CheckDlgButton(IDC_CHECKTRACK+IDC_CHECKNOTRACK-wID,BST_UNCHECKED);
	return 0;
}

LRESULT CEditMenuDlg::OnCheckMulti( WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled )
{
	if (IsDlgButtonChecked(IDC_CHECKMULTICOLUMN)==BST_CHECKED)
		m_pItem->settings|=StdMenuItem::MENU_MULTICOLUMN;
	else
		m_pItem->settings&=~StdMenuItem::MENU_MULTICOLUMN;
	return 0;
}

LRESULT CEditMenuDlg::OnBrowseCommand( WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled )
{
	wchar_t text[_MAX_PATH];
	GetDlgItemText(IDC_COMBOCOMMAND,text,_countof(text));
	if (BrowseCommand(m_hWnd,text))
	{
		SetDlgItemText(IDC_COMBOCOMMAND,text);
		SendMessage(WM_COMMAND,MAKEWPARAM(IDC_COMBOCOMMAND,CBN_KILLFOCUS));
	}
	return 0;
}

LRESULT CEditMenuDlg::OnBrowseLink( WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled )
{
	wchar_t text[_MAX_PATH];
	GetDlgItemText(IDC_COMBOLINK,text,_countof(text));
	if (BrowseLink(m_hWnd,text))
	{
		SetDlgItemText(IDC_COMBOLINK,text);
		SendMessage(WM_COMMAND,MAKEWPARAM(IDC_COMBOLINK,CBN_KILLFOCUS));
	}
	return 0;
}

LRESULT CEditMenuDlg::OnBrowseIcon( WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled )
{
	wchar_t text[_MAX_PATH];
	GetDlgItemText(IDC_EDITICON,text,_countof(text));
	if (BrowseIcon(text))
	{
		SetDlgItemText(IDC_EDITICON,text);
		SendMessage(WM_COMMAND,MAKEWPARAM(IDC_EDITICON,EN_KILLFOCUS));
	}
	return 0;
}

LRESULT CEditMenuDlg::OnReset( WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled )
{
	if (!m_pItem->pStdCommand || !*m_pItem->pStdCommand->name)
		return 0;
	m_pItem->link.Empty();
	m_pItem->label=m_pItem->pStdCommand->label;
	m_pItem->tip=m_pItem->pStdCommand->tip;
	m_pItem->icon=m_pItem->pStdCommand->icon;
	m_pItem->iconD=m_pItem->pStdCommand->iconD;
	m_pItem->settings=m_pItem->pStdCommand->settings;
	SendDlgItemMessage(IDC_COMBOLINK,CB_SETCURSEL,-1);
	SetDlgItemText(IDC_EDITLABEL,m_pItem->label);
	SetDlgItemText(IDC_EDITTIP,m_pItem->tip);
	SetDlgItemText(IDC_EDITICON,m_pItem->icon);
	CheckDlgButton(IDC_CHECKSORTZA,(m_pItem->settings&StdMenuItem::MENU_SORTZA)?BST_CHECKED:BST_UNCHECKED);
	CheckDlgButton(IDC_CHECKSORTZAREC,(m_pItem->settings&StdMenuItem::MENU_SORTZA_REC)?BST_CHECKED:BST_UNCHECKED);
	CheckDlgButton(IDC_CHECKSORTONCE,(m_pItem->settings&StdMenuItem::MENU_SORTONCE)?BST_CHECKED:BST_UNCHECKED);
	CheckDlgButton(IDC_CHECKOPENUP,(m_pItem->settings&StdMenuItem::MENU_OPENUP)?BST_CHECKED:BST_UNCHECKED);
	CheckDlgButton(IDC_CHECKOPENUPREC,(m_pItem->settings&StdMenuItem::MENU_OPENUP_REC)?BST_CHECKED:BST_UNCHECKED);
	CheckDlgButton(IDC_CHECKNOEXPAND,(m_pItem->settings&StdMenuItem::MENU_NOEXPAND)?BST_CHECKED:BST_UNCHECKED);
	CheckDlgButton(IDC_CHECKMULTICOLUMN,(m_pItem->settings&StdMenuItem::MENU_MULTICOLUMN)?BST_CHECKED:BST_UNCHECKED);
	CheckDlgButton(IDC_CHECKTRACK,(m_pItem->settings&StdMenuItem::MENU_TRACK)?BST_CHECKED:BST_UNCHECKED);
	CheckDlgButton(IDC_CHECKNOTRACK,(m_pItem->settings&StdMenuItem::MENU_NOTRACK)?BST_CHECKED:BST_UNCHECKED);
	CheckDlgButton(IDC_CHECKITEMSFIRST,(m_pItem->settings&StdMenuItem::MENU_ITEMS_FIRST)?BST_CHECKED:BST_UNCHECKED);
	CheckDlgButton(IDC_CHECKINLINE,(m_pItem->settings&StdMenuItem::MENU_INLINE)?BST_CHECKED:BST_UNCHECKED);
	CheckDlgButton(IDC_CHECKNOEXT,(m_pItem->settings&StdMenuItem::MENU_NOEXTENSIONS)?BST_CHECKED:BST_UNCHECKED);
	if (m_pItem->pStdCommand && wcscmp(m_pItem->pStdCommand->name,L"programs")==0)
	{
		CheckDlgButton(IDC_CHECKMULTICOLUMN,BST_CHECKED);
		GetDlgItem(IDC_CHECKMULTICOLUMN).EnableWindow(FALSE);
	}
	else
		GetDlgItem(IDC_CHECKMULTICOLUMN).EnableWindow(FALSE);

	UpdateIcons(IDC_ICONN,IDC_ICOND);
	return 0;
}

///////////////////////////////////////////////////////////////////////////////

class CCustomMenuDlg: public CCustomTreeDlg
{
public:
	CCustomMenuDlg( void ): CCustomTreeDlg(true,g_StdCommands) {}
	bool Validate( void );

protected:
	virtual void ParseTreeItemExtra( CTreeItem *pItem, CSettingsParser &parser );
	virtual void SerializeItemExtra( CTreeItem *pItem, std::vector<wchar_t> &stringBuilder );
	virtual bool EditItem( CTreeItem *pItem, HWND tree, HTREEITEM hItem, std::vector<HMODULE> &modules );
	virtual void InitItems( void ) { UpdateWarnings(); }
	virtual void ItemsChanged( void ) { UpdateWarnings(); }

private:
	void UpdateWarnings( void );
	bool FindMenuItem( HTREEITEM hParent, const wchar_t *command );
};

void CCustomMenuDlg::ParseTreeItemExtra( CTreeItem *pItem, CSettingsParser &parser )
{
	pItem->settings=0;
	wchar_t text[256];
	Sprintf(text,_countof(text),L"%s.Settings",pItem->name);
	const wchar_t *str=parser.FindSetting(text);
	if (!str) return;

	while(*str)
	{
		wchar_t token[256];
		str=GetToken(str,token,_countof(token),L", \t|;");
		if (_wcsicmp(token,L"OPEN_UP")==0) pItem->settings|=StdMenuItem::MENU_OPENUP;
		if (_wcsicmp(token,L"OPEN_UP_CHILDREN")==0) pItem->settings|=StdMenuItem::MENU_OPENUP_REC;
		if (_wcsicmp(token,L"SORT_ZA")==0) pItem->settings|=StdMenuItem::MENU_SORTZA;
		if (_wcsicmp(token,L"SORT_ZA_CHILDREN")==0) pItem->settings|=StdMenuItem::MENU_SORTZA_REC;
		if (_wcsicmp(token,L"SORT_ONCE")==0) pItem->settings|=StdMenuItem::MENU_SORTONCE;
		if (_wcsicmp(token,L"ITEMS_FIRST")==0) pItem->settings|=StdMenuItem::MENU_ITEMS_FIRST;
		if (_wcsicmp(token,L"TRACK_RECENT")==0) pItem->settings|=StdMenuItem::MENU_TRACK;
		if (_wcsicmp(token,L"NOTRACK_RECENT")==0) pItem->settings|=StdMenuItem::MENU_NOTRACK;
		if (_wcsicmp(token,L"NOEXPAND")==0) pItem->settings|=StdMenuItem::MENU_NOEXPAND;
		if (_wcsicmp(token,L"MULTICOLUMN")==0) pItem->settings|=StdMenuItem::MENU_MULTICOLUMN;
		if (_wcsicmp(token,L"INLINE")==0) pItem->settings|=StdMenuItem::MENU_INLINE;
		if (_wcsicmp(token,L"NOEXTENSIONS")==0) pItem->settings|=StdMenuItem::MENU_NOEXTENSIONS;
	}
}

void CCustomMenuDlg::SerializeItemExtra( CTreeItem *pItem, std::vector<wchar_t> &stringBuilder )
{
	if (!pItem->settings) return;
	wchar_t text[256];
	Sprintf(text,_countof(text),L"%s.Settings=",pItem->name);
	AppendString(stringBuilder,text);
	if (pItem->settings&StdMenuItem::MENU_OPENUP) AppendString(stringBuilder,L"OPEN_UP|");
	if (pItem->settings&StdMenuItem::MENU_OPENUP_REC) AppendString(stringBuilder,L"OPEN_UP_CHILDREN|");
	if (pItem->settings&StdMenuItem::MENU_SORTZA) AppendString(stringBuilder,L"SORT_ZA|");
	if (pItem->settings&StdMenuItem::MENU_SORTZA_REC) AppendString(stringBuilder,L"SORT_ZA_CHILDREN|");
	if (pItem->settings&StdMenuItem::MENU_SORTONCE) AppendString(stringBuilder,L"SORT_ONCE|");
	if (pItem->settings&StdMenuItem::MENU_ITEMS_FIRST) AppendString(stringBuilder,L"ITEMS_FIRST|");
	if (pItem->settings&StdMenuItem::MENU_TRACK) AppendString(stringBuilder,L"TRACK_RECENT|");
	if (pItem->settings&StdMenuItem::MENU_NOTRACK) AppendString(stringBuilder,L"NOTRACK_RECENT|");
	if (pItem->settings&StdMenuItem::MENU_NOEXPAND) AppendString(stringBuilder,L"NOEXPAND|");
	if (pItem->settings&StdMenuItem::MENU_MULTICOLUMN) AppendString(stringBuilder,L"MULTICOLUMN|");
	if (pItem->settings&StdMenuItem::MENU_INLINE) AppendString(stringBuilder,L"INLINE|");
	if (pItem->settings&StdMenuItem::MENU_NOEXTENSIONS) AppendString(stringBuilder,L"NOEXTENSIONS|");
	stringBuilder[stringBuilder.size()-1]='\n';
}

bool CCustomMenuDlg::EditItem( CTreeItem *pItem, HWND tree, HTREEITEM hItem, std::vector<HMODULE> &modules )
{
	return CEditMenuDlg(pItem,modules).Run(m_hWnd,IDD_CUSTOMMENU);
}

void CCustomMenuDlg::UpdateWarnings( void )
{
	CSettingsLockWrite lock;
	bool bWarning;
	bWarning=!FindMenuItem(NULL,L"favorites");
	UpdateSetting(L"Favorites",bWarning?IDS_SHOW_FAVORITES_TIP2:IDS_SHOW_FAVORITES_TIP,bWarning);

	bWarning=!FindMenuItem(NULL,L"documents");
	UpdateSetting(L"Documents",bWarning?IDS_SHOW_DOCUMENTS_TIP2:IDS_SHOW_DOCUMENTS_TIP,bWarning);

	bWarning=!FindMenuItem(NULL,L"user_files");
	UpdateSetting(L"UserFiles",bWarning?IDS_SHOW_USERFILES_TIP2:IDS_SHOW_USERFILES_TIP,bWarning);

	bWarning=!FindMenuItem(NULL,L"user_documents");
	UpdateSetting(L"UserDocuments",bWarning?IDS_SHOW_USERDOCS_TIP2:IDS_SHOW_USERDOCS_TIP,bWarning);

	bWarning=!FindMenuItem(NULL,L"user_pictures");
	UpdateSetting(L"UserPictures",bWarning?IDS_SHOW_USERPICS_TIP2:IDS_SHOW_USERPICS_TIP,bWarning);

	bWarning=!FindMenuItem(NULL,L"control_panel") && !FindMenuItem(NULL,L"control_panel_categories");
	UpdateSetting(L"ControlPanel",bWarning?IDS_SHOW_CP_TIP2:IDS_SHOW_CP_TIP,bWarning);

	bWarning=!FindMenuItem(NULL,L"network_connections");
	UpdateSetting(L"Network",bWarning?IDS_SHOW_NETWORK_TIP2:IDS_SHOW_NETWORK_TIP,bWarning);

	bWarning=!FindMenuItem(NULL,L"printers");
	UpdateSetting(L"Printers",bWarning?IDS_SHOW_PRINTERS_TIP2:IDS_SHOW_PRINTERS_TIP,bWarning);

	bWarning=!FindMenuItem(NULL,L"shutdown_box");
	UpdateSetting(L"Shutdown",bWarning?IDS_SHOW_SHUTDOWN_TIP2:IDS_SHOW_SHUTDOWN_TIP,bWarning);

	bWarning=!FindMenuItem(NULL,L"search_box");
	UpdateSetting(L"SearchBox",bWarning?IDS_SHOW_SEARCH_BOX_TIP2:IDS_SHOW_SEARCH_BOX_TIP,bWarning);

	bWarning=!FindMenuItem(NULL,L"search");
	UpdateSetting(L"Search",bWarning?IDS_SHOW_SEARCH_TIP2:IDS_SHOW_SEARCH_TIP,bWarning);

	bWarning=!FindMenuItem(NULL,L"help");
	UpdateSetting(L"Help",bWarning?IDS_SHOW_HELP_TIP2:IDS_SHOW_HELP_TIP,bWarning);

	bWarning=!FindMenuItem(NULL,L"run");
	UpdateSetting(L"Run",bWarning?IDS_SHOW_RUN_TIP2:IDS_SHOW_RUN_TIP,bWarning);

	bWarning=!FindMenuItem(NULL,L"logoff");
	UpdateSetting(L"LogOff",bWarning?IDS_SHOW_LOGOFF_TIP2:IDS_SHOW_LOGOFF_TIP,bWarning);

	bWarning=!FindMenuItem(NULL,L"undock");
	UpdateSetting(L"Undock",bWarning?IDS_SHOW_UNDOCK_TIP2:IDS_SHOW_UNDOCK_TIP,bWarning);
}

bool CCustomMenuDlg::FindMenuItem( HTREEITEM hParent, const wchar_t *command )
{
	if (hParent)
	{
		CTreeItem *pItem=GetItem(hParent);
		if (!pItem) return false;
		if (_wcsicmp(pItem->command,command)==0)
			return true;
	}
	for (HTREEITEM hChild=hParent?GetChild(hParent):GetRoot();hChild;hChild=GetNext(hChild))
		if (FindMenuItem(hChild,command))
			return true;
	return false;
}

class CCustomMenuPanel: public ISettingsPanel
{
public:
	virtual HWND Create( HWND parent );
	virtual HWND Activate( CSetting *pGroup, const RECT &rect, bool bReset );
	virtual bool Validate( HWND parent ) { return true; }

private:
	static CCustomMenuDlg s_Dialog;
};

CCustomMenuDlg CCustomMenuPanel::s_Dialog;

HWND CCustomMenuPanel::Create( HWND parent )
{
	if (!s_Dialog.m_hWnd)
		s_Dialog.Create(parent,LoadDialogEx(IDD_CUSTOMTREE));
	return s_Dialog.m_hWnd;
}

HWND CCustomMenuPanel::Activate( CSetting *pGroup, const RECT &rect, bool bReset )
{
	s_Dialog.SetGroup(pGroup,bReset);
	s_Dialog.SetWindowPos(HWND_TOP,&rect,SWP_SHOWWINDOW);
	return s_Dialog.m_hWnd;
}

static CCustomMenuPanel g_CustomMenuPanel;

///////////////////////////////////////////////////////////////////////////////

CSetting g_Settings[]={
{L"Basic",CSetting::TYPE_GROUP,IDS_BASIC_SETTINGS},
	{L"EnableSettings",CSetting::TYPE_BOOL,0,0,1,CSetting::FLAG_HIDDEN},
	{L"CrashDump",CSetting::TYPE_INT,0,0,0,CSetting::FLAG_HIDDEN},
	{L"LogLevel",CSetting::TYPE_INT,0,0,0,CSetting::FLAG_HIDDEN},

{L"Controls",CSetting::TYPE_GROUP,IDS_CONTROLS_SETTINGS},
	{L"MouseClick",CSetting::TYPE_INT,IDS_LCLICK,IDS_LCLICK_TIP,1,CSetting::FLAG_BASIC},
		{L"Nothing",CSetting::TYPE_RADIO,IDS_OPEN_NOTHING,IDS_OPEN_NOTHING_TIP},
		{L"ClassicMenu",CSetting::TYPE_RADIO,IDS_OPEN_CSM,IDS_OPEN_CSM_TIP},
		{L"WindowsMenu",CSetting::TYPE_RADIO,IDS_OPEN_WSM,IDS_OPEN_WSM_TIP},
	{L"ShiftClick",CSetting::TYPE_INT,IDS_SHIFT_LCLICK,IDS_SHIFT_LCLICK_TIP,2,CSetting::FLAG_BASIC},
		{L"Nothing",CSetting::TYPE_RADIO,IDS_OPEN_NOTHING,IDS_OPEN_NOTHING_TIP},
		{L"ClassicMenu",CSetting::TYPE_RADIO,IDS_OPEN_CSM,IDS_OPEN_CSM_TIP},
		{L"WindowsMenu",CSetting::TYPE_RADIO,IDS_OPEN_WSM,IDS_OPEN_WSM_TIP},
	{L"WinKey",CSetting::TYPE_INT,IDS_WIN_KEY,IDS_WIN_KEY_TIP,1,CSetting::FLAG_BASIC},
		{L"Nothing",CSetting::TYPE_RADIO,IDS_OPEN_NOTHING,IDS_OPEN_NOTHING_TIP},
		{L"ClassicMenu",CSetting::TYPE_RADIO,IDS_OPEN_CSM,IDS_OPEN_CSM_TIP},
		{L"WindowsMenu",CSetting::TYPE_RADIO,IDS_OPEN_WSM,IDS_OPEN_WSM_TIP},
	{L"ShiftWin",CSetting::TYPE_INT,IDS_SHIFT_WIN,IDS_SHIFT_WIN_TIP,2,CSetting::FLAG_BASIC},
		{L"Nothing",CSetting::TYPE_RADIO,IDS_OPEN_NOTHING,IDS_OPEN_NOTHING_TIP},
		{L"ClassicMenu",CSetting::TYPE_RADIO,IDS_OPEN_CSM,IDS_OPEN_CSM_TIP},
		{L"WindowsMenu",CSetting::TYPE_RADIO,IDS_OPEN_WSM,IDS_OPEN_WSM_TIP},
	{L"MiddleClick",CSetting::TYPE_INT,IDS_MCLICK,IDS_MCLICK_TIP,0},
		{L"Nothing",CSetting::TYPE_RADIO,IDS_OPEN_NOTHING,IDS_OPEN_NOTHING_TIP},
		{L"ClassicMenu",CSetting::TYPE_RADIO,IDS_OPEN_CSM,IDS_OPEN_CSM_TIP},
		{L"WindowsMenu",CSetting::TYPE_RADIO,IDS_OPEN_WSM,IDS_OPEN_WSM_TIP},
	{L"Hover",CSetting::TYPE_INT,IDS_HOVER,IDS_HOVER_TIP,0},
		{L"Nothing",CSetting::TYPE_RADIO,IDS_OPEN_NOTHING,IDS_OPEN_NOTHING_TIP},
		{L"ClassicMenu",CSetting::TYPE_RADIO,IDS_OPEN_CSM,IDS_OPEN_CSM_TIP},
		{L"WindowsMenu",CSetting::TYPE_RADIO,IDS_OPEN_WSM,IDS_OPEN_WSM_TIP},
	{L"StartHoverDelay",CSetting::TYPE_INT,IDS_HOVER_DELAY,IDS_HOVER_DELAY_TIP,1000,0,L"Hover"},
	{L"CSMHotkey",CSetting::TYPE_HOTKEY,IDS_CSM_HOTKEY,IDS_CSM_HOTKEY_TIP,0},
	{L"WSMHotkey",CSetting::TYPE_HOTKEY,IDS_WSM_HOTKEY,IDS_WSM_HOTKEY_TIP,0},

{L"SpecialItems",CSetting::TYPE_GROUP,IDS_SHOW_ITEMS},
	{L"Favorites",CSetting::TYPE_INT,IDS_SHOW_FAVORITES,IDS_SHOW_FAVORITES_TIP,0,CSetting::FLAG_BASIC},
		{L"Hide",CSetting::TYPE_RADIO,IDS_ITEM_HIDE,IDS_ITEM_HIDE_TIP},
		{L"Show",CSetting::TYPE_RADIO,IDS_ITEM_SHOW,IDS_ITEM_SHOW_TIP},
		{L"Menu",CSetting::TYPE_RADIO,IDS_ITEM_MENU,IDS_ITEM_MENU_TIP},
	{L"Documents",CSetting::TYPE_INT,IDS_SHOW_DOCUMENTS,IDS_SHOW_DOCUMENTS_TIP,2,CSetting::FLAG_BASIC},
		{L"Hide",CSetting::TYPE_RADIO,IDS_ITEM_HIDE,IDS_ITEM_HIDE_TIP},
		{L"Show",CSetting::TYPE_RADIO,IDS_ITEM_SHOW,IDS_ITEM_SHOW_TIP},
		{L"Menu",CSetting::TYPE_RADIO,IDS_ITEM_MENU,IDS_ITEM_MENU_TIP},
	{L"MaxRecentDocuments",CSetting::TYPE_INT,IDS_MAX_DOCS,IDS_MAX_DOCS_TIP,15,0,L"Documents=2"},
	{L"UserFiles",CSetting::TYPE_INT,IDS_SHOW_USERFILES,IDS_SHOW_USERFILES_TIP,1},
		{L"Hide",CSetting::TYPE_RADIO,IDS_ITEM_HIDE,IDS_ITEM_HIDE_TIP},
		{L"Show",CSetting::TYPE_RADIO,IDS_ITEM_SHOW,IDS_ITEM_SHOW_TIP},
		{L"Menu",CSetting::TYPE_RADIO,IDS_ITEM_MENU,IDS_ITEM_MENU_TIP},
	{L"UserDocuments",CSetting::TYPE_INT,IDS_SHOW_USERDOCS,IDS_SHOW_USERDOCS_TIP,1},
		{L"Hide",CSetting::TYPE_RADIO,IDS_ITEM_HIDE,IDS_ITEM_HIDE_TIP},
		{L"Show",CSetting::TYPE_RADIO,IDS_ITEM_SHOW,IDS_ITEM_SHOW_TIP},
		{L"Menu",CSetting::TYPE_RADIO,IDS_ITEM_MENU,IDS_ITEM_MENU_TIP},
	{L"UserPictures",CSetting::TYPE_INT,IDS_SHOW_USERPICS,IDS_SHOW_USERPICS_TIP,1},
		{L"Hide",CSetting::TYPE_RADIO,IDS_ITEM_HIDE,IDS_ITEM_HIDE_TIP},
		{L"Show",CSetting::TYPE_RADIO,IDS_ITEM_SHOW,IDS_ITEM_SHOW_TIP},
		{L"Menu",CSetting::TYPE_RADIO,IDS_ITEM_MENU,IDS_ITEM_MENU_TIP},
	{L"ControlPanel",CSetting::TYPE_INT,IDS_SHOW_CP,IDS_SHOW_CP_TIP,1,CSetting::FLAG_BASIC},
		{L"Hide",CSetting::TYPE_RADIO,IDS_ITEM_HIDE,IDS_ITEM_HIDE_TIP},
		{L"Show",CSetting::TYPE_RADIO,IDS_ITEM_SHOW,IDS_ITEM_SHOW_TIP},
		{L"Menu",CSetting::TYPE_RADIO,IDS_ITEM_MENU,IDS_ITEM_MENU_TIP},
	{L"Network",CSetting::TYPE_INT,IDS_SHOW_NETWORK,IDS_SHOW_NETWORK_TIP,1},
		{L"Hide",CSetting::TYPE_RADIO,IDS_ITEM_HIDE,IDS_ITEM_HIDE_TIP},
		{L"Show",CSetting::TYPE_RADIO,IDS_ITEM_SHOW,IDS_ITEM_SHOW_TIP},
		{L"Menu",CSetting::TYPE_RADIO,IDS_ITEM_MENU,IDS_ITEM_MENU_TIP},
	{L"Printers",CSetting::TYPE_INT,IDS_SHOW_PRINTERS,IDS_SHOW_PRINTERS_TIP,1},
		{L"Hide",CSetting::TYPE_RADIO,IDS_ITEM_HIDE,IDS_ITEM_HIDE_TIP},
		{L"Show",CSetting::TYPE_RADIO,IDS_ITEM_SHOW,IDS_ITEM_SHOW_TIP},
		{L"Menu",CSetting::TYPE_RADIO,IDS_ITEM_MENU,IDS_ITEM_MENU_TIP},
	{L"Shutdown",CSetting::TYPE_INT,IDS_SHOW_SHUTDOWN,IDS_SHOW_SHUTDOWN_TIP,1},
		{L"Hide",CSetting::TYPE_RADIO,IDS_ITEM_HIDE,IDS_ITEM_HIDE_TIP},
		{L"Show",CSetting::TYPE_RADIO,IDS_ITEM_SHOW,IDS_ITEM_SHOW_TIP},
		{L"Menu",CSetting::TYPE_RADIO,IDS_ITEM_MENU,IDS_ITEM_MENU_TIP},
	{L"Search",CSetting::TYPE_BOOL,IDS_SHOW_SEARCH,IDS_SHOW_SEARCH_TIP,1},
	{L"Help",CSetting::TYPE_BOOL,IDS_SHOW_HELP,IDS_SHOW_HELP_TIP,1},
	{L"Run",CSetting::TYPE_BOOL,IDS_SHOW_RUN,IDS_SHOW_RUN_TIP,1},
	{L"LogOff",CSetting::TYPE_BOOL,IDS_SHOW_LOGOFF,IDS_SHOW_LOGOFF_TIP,0,CSetting::FLAG_BASIC},
	{L"ConfirmLogOff",CSetting::TYPE_BOOL,IDS_CONFIRM_LOGOFF,IDS_CONFIRM_LOGOFF_TIP,0,0,L"LogOff"},
	{L"Undock",CSetting::TYPE_BOOL,IDS_SHOW_UNDOCK,IDS_SHOW_UNDOCK_TIP,1},
	{L"RemoteShutdown",CSetting::TYPE_BOOL,IDS_SHOW_RSHUTDOWN,IDS_SHOW_RSHUTDOWN_TIP,0},
	{L"RecentPrograms",CSetting::TYPE_BOOL,IDS_SHOW_RECENT,IDS_SHOW_RECENT_TIP,0,CSetting::FLAG_BASIC},
	{L"MaxRecentPrograms",CSetting::TYPE_INT,IDS_MAX_PROGS,IDS_MAX_PROGS_TIP,5,0,L"RecentPrograms"},
	{L"RecentProgsTop",CSetting::TYPE_BOOL,IDS_RECENT_TOP,IDS_RECENT_TOP_TIP,1,0,L"RecentPrograms"},
	{L"RecentProgKeys",CSetting::TYPE_INT,IDS_RECENT_KEYS,IDS_RECENT_KEYS_TIP,2,0,L"RecentPrograms"},
		{L"NoKey",CSetting::TYPE_RADIO,IDS_KEY_NOTHING,IDS_KEY_NOTHING_TIP,0,0,L"RecentPrograms"},
		{L"Normal",CSetting::TYPE_RADIO,IDS_KEY_NORMAL,IDS_KEY_NORMAL_TIP,0,0,L"RecentPrograms"},
		{L"Digits",CSetting::TYPE_RADIO,IDS_KEY_DIGITS,IDS_KEY_DIGITS_TIP,0,0,L"RecentPrograms"},
		{L"HiddenDigits",CSetting::TYPE_RADIO,IDS_KEY_HIDDEN,IDS_KEY_HIDDEN_TIP,0,0,L"RecentPrograms"},

{L"GeneralBehavior",CSetting::TYPE_GROUP,IDS_BEHAVIOR_SETTINGS},
	{L"ExpandFolderLinks",CSetting::TYPE_BOOL,IDS_EXPAND_LINKS,IDS_EXPAND_LINKS_TIP,1},
	{L"MenuDelay",CSetting::TYPE_INT,IDS_MENU_DELAY,IDS_MENU_DELAY_TIP,-1,CSetting::FLAG_BASIC}, // system delay time
	{L"InfotipDelay",CSetting::TYPE_STRING,IDS_TIP_DELAY,IDS_TIP_DELAY_TIP,L"400,4000"},
	{L"FolderInfotipDelay",CSetting::TYPE_STRING,IDS_FTIP_DELAY,IDS_FTIP_DELAY_TIP,L"0,0"},
	{L"MainMenuAnimation",CSetting::TYPE_INT,IDS_ANIMATION,IDS_ANIMATION_TIP,-1}, // system animation type
		{L"None",CSetting::TYPE_RADIO,IDS_ANIM_NONE,IDS_ANIM_NONE_TIP},
		{L"Fade",CSetting::TYPE_RADIO,IDS_ANIM_FADE,IDS_ANIM_FADE_TIP},
		{L"Slide",CSetting::TYPE_RADIO,IDS_ANIM_SLIDE,IDS_ANIM_SLIDE_TIP},
		{L"Random",CSetting::TYPE_RADIO,IDS_ANIM_RANDOM,IDS_ANIM_RANDOM_TIP},
	{L"MainMenuAnimationSpeed",CSetting::TYPE_INT,IDS_ANIM_SPEED,IDS_ANIM_SPEED_TIP,200,0,L"MainMenuAnimation"},
	{L"SubMenuAnimation",CSetting::TYPE_INT,IDS_SUB_ANIMATION,IDS_SUB_ANIMATION_TIP,-1}, // system animation type
		{L"None",CSetting::TYPE_RADIO,IDS_ANIM_NONE,IDS_ANIM_NONE_TIP},
		{L"Fade",CSetting::TYPE_RADIO,IDS_ANIM_FADE,IDS_ANIM_FADE_TIP},
		{L"Slide",CSetting::TYPE_RADIO,IDS_ANIM_SLIDE,IDS_ANIM_SLIDE_TIP},
		{L"Random",CSetting::TYPE_RADIO,IDS_ANIM_RANDOM,IDS_ANIM_RANDOM_TIP},
	{L"SubMenuAnimationSpeed",CSetting::TYPE_INT,IDS_SUB_ANIM_SPEED,IDS_SUB_ANIM_SPEED_TIP,200,0,L"SubMenuAnimation"},
	{L"MainMenuScrollSpeed",CSetting::TYPE_INT,IDS_SCROLL_SPEED,IDS_SCROLL_SPEED_TIP,3},
	{L"SubMenuScrollSpeed",CSetting::TYPE_INT,IDS_SUB_SCROLL_SPEED,IDS_SUB_SCROLL_SPEED_TIP,3},
	{L"MenuFadeSpeed",CSetting::TYPE_INT,IDS_FADE_SPEED,IDS_FADE_SPEED_TIP,400},
	{L"DragHideDelay",CSetting::TYPE_INT,IDS_DRAG_DELAY,IDS_DRAG_DELAY_TIP,4000},
	{L"EnableAccessibility",CSetting::TYPE_BOOL,IDS_ACCESSIBILITY,IDS_ACCESSIBILITY_TIP,1},
	{L"ShowNextToTaskbar",CSetting::TYPE_BOOL,IDS_NEXTTASKBAR,IDS_NEXTTASKBAR_TIP,0},
	{L"UserPictureCommand",CSetting::TYPE_STRING,IDS_PIC_COMMAND,IDS_PIC_COMMAND_TIP,L"control nusrmgr.cpl"},
	{L"UserNameCommand",CSetting::TYPE_STRING,IDS_NAME_COMMAND,IDS_NAME_COMMAND_TIP,L"control nusrmgr.cpl"},
	{L"SearchFilesCommand",CSetting::TYPE_STRING,IDS_SEARCH_COMMAND,IDS_SEARCH_COMMAND_TIP,L""},
	{L"CascadingMenu",CSetting::TYPE_BOOL,IDS_CASCADE_MENU,IDS_CASCADE_MENU_TIP,0},
	{L"MainSortZA",CSetting::TYPE_BOOL,IDS_MAIN_SORTZA,IDS_SORTZA_TIP,0},
	{L"MainSortOnce",CSetting::TYPE_BOOL,IDS_MAIN_SORTONCE,IDS_SORTONCE_TIP,0},
	{L"PreCacheIcons",CSetting::TYPE_BOOL,IDS_CACHE_ICONS,IDS_CACHE_ICONS_TIP,1,CSetting::FLAG_COLD},
	{L"DelayIcons",CSetting::TYPE_BOOL,IDS_DELAY_ICONS,IDS_DELAY_ICONS_TIP,1,CSetting::FLAG_COLD},
	{L"ReportSkinErrors",CSetting::TYPE_BOOL,IDS_SKIN_ERRORS,IDS_SKIN_ERRORS_TIP,0},

{L"Look",CSetting::TYPE_GROUP,IDS_LOOK_SETTINGS},
	{L"ScrollType",CSetting::TYPE_INT,IDS_SCROLL_TYPE,IDS_SCROLL_TYPE_TIP,1},
		{L"Scroll",CSetting::TYPE_RADIO,IDS_SCROLL_YES,IDS_SCROLL_YES_TIP},
		{L"NoScroll",CSetting::TYPE_RADIO,IDS_SCROLL_NO,IDS_SCROLL_NO_TIP},
		{L"Auto",CSetting::TYPE_RADIO,IDS_SCROLL_AUTO,IDS_SCROLL_AUTO_TIP},
	{L"SameSizeColumns",CSetting::TYPE_BOOL,IDS_SAME_COLUMNS,IDS_SAME_COLUMNS_TIP,1,0,L"ScrollType"},
	{L"MaxMainMenuWidth",CSetting::TYPE_INT,IDS_MENU_WIDTH,IDS_MENU_WIDTH_TIP,60},
	{L"MaxMenuWidth",CSetting::TYPE_INT,IDS_SUBMENU_WIDTH,IDS_SUBMENU_WIDTH_TIP,60},
	{L"MenuCaption",CSetting::TYPE_STRING,IDS_MENU_CAPTION,IDS_MENU_CAPTION_TIP,L""},
	{L"MenuUsername",CSetting::TYPE_STRING,IDS_MENU_USERNAME,IDS_MENU_USERNAME_TIP,L""},
	{L"SmallIconSize",CSetting::TYPE_INT,IDS_SMALL_SIZE_SM,IDS_SMALL_SIZE_SM_TIP,-1,CSetting::FLAG_COLD}, // 16 for DPI<=96, 20 for DPI<=120, 24 otherwise
	{L"LargeIconSize",CSetting::TYPE_INT,IDS_LARGE_SIZE_SM,IDS_LARGE_SIZE_SM_TIP,-1,CSetting::FLAG_COLD}, // 32 for DPI<=96, 40 for DPI<=120, 48 otherwise
	{L"NumericSort",CSetting::TYPE_BOOL,IDS_NUMERIC_SORT,IDS_NUMERIC_SORT_TIP,1},
	{L"FontSmoothing",CSetting::TYPE_INT,IDS_FONT_SMOOTHING,IDS_FONT_SMOOTHING_TIP,0},
		{L"Default",CSetting::TYPE_RADIO,IDS_SMOOTH_DEFAULT,IDS_SMOOTH_DEFAULT_TIP},
		{L"None",CSetting::TYPE_RADIO,IDS_SMOOTH_NONE,IDS_SMOOTH_NONE_TIP},
		{L"Standard",CSetting::TYPE_RADIO,IDS_SMOOTH_STD,IDS_SMOOTH_STD_TIP},
		{L"ClearType",CSetting::TYPE_RADIO,IDS_SMOOTH_CLEAR,IDS_SMOOTH_CLEAR_TIP},

{L"SearchBoxSettings",CSetting::TYPE_GROUP,IDS_SEARCH_BOX},
	{L"SearchBox",CSetting::TYPE_INT,IDS_SHOW_SEARCH_BOX,IDS_SHOW_SEARCH_BOX_TIP,0,CSetting::FLAG_BASIC},
		{L"Hide",CSetting::TYPE_RADIO,IDS_SEARCH_BOX_HIDE,IDS_SEARCH_BOX_HIDE_TIP},
		{L"Normal",CSetting::TYPE_RADIO,IDS_SEARCH_BOX_SHOW,IDS_SEARCH_BOX_SHOW_TIP},
		{L"Tab",CSetting::TYPE_RADIO,IDS_SEARCH_BOX_TAB,IDS_SEARCH_BOX_TAB_TIP},
	{L"SearchSelect",CSetting::TYPE_BOOL,IDS_SEARCH_BOX_SEL,IDS_SEARCH_BOX_SEL_TIP,1,0,L"SearchBox=1"},
	{L"SearchCP",CSetting::TYPE_BOOL,IDS_SEARCH_CP,IDS_SEARCH_CP_TIP,1,0,L"SearchBox"},
	{L"SearchPath",CSetting::TYPE_BOOL,IDS_SEARCH_PATH,IDS_SEARCH_PATH_TIP,1,0,L"SearchBox"},
	{L"SearchSubWord",CSetting::TYPE_BOOL,IDS_SUB_WORD,IDS_SUB_WORD_TIP,1,0,L"SearchBox"},
	{L"SearchTrack",CSetting::TYPE_BOOL,IDS_SEARCH_TRACK,IDS_SEARCH_TRACK_TIP,1,0,L"SearchBox"},
	{L"SearchMax",CSetting::TYPE_INT,IDS_SEARCH_MAX,IDS_SEARCH_MAX_TIP,20,0,L"SearchBox"},
	{L"SearchAutoComplete",CSetting::TYPE_BOOL,IDS_SEARCH_AUTO,IDS_SEARCH_AUTO_TIP,1,0,L"SearchBox"},

{L"Skin",CSetting::TYPE_GROUP,IDS_SKIN_SETTINGS,0,0,CSetting::FLAG_BASIC,NULL,&g_SkinSettingsPanel},
	{L"Skin1",CSetting::TYPE_STRING,0,0,L"Windows XP Luna"},
	{L"SkinVariation1",CSetting::TYPE_STRING,0,0,L""},
	{L"SkinOptions1",CSetting::TYPE_STRING,0,0,L""},

{L"Security",CSetting::TYPE_GROUP,IDS_SECURITY_SETTINGS},
	{L"EnableDragDrop",CSetting::TYPE_BOOL,IDS_DRAG_DROP,IDS_DRAG_DROP_TIP,1},
	{L"EnableContextMenu",CSetting::TYPE_BOOL,IDS_CONTEXT_MENU,IDS_CONTEXT_MENU_TIP,1},
	{L"ShowNewFolder",CSetting::TYPE_BOOL,IDS_NEW_FOLDER,IDS_NEW_FOLDER_TIP,1,0,L"EnableContextMenu"},
	{L"EnableExit",CSetting::TYPE_BOOL,IDS_EXIT,IDS_EXIT_TIP,1},

{L"Sounds",CSetting::TYPE_GROUP,IDS_SOUND_SETTINGS},
	{L"SoundMain",CSetting::TYPE_SOUND,IDS_SOUND_MAIN,IDS_SOUND_MAIN_TIP,L"MenuPopup"},
	{L"SoundPopup",CSetting::TYPE_SOUND,IDS_SOUND_POPUP,IDS_SOUND_POPUP_TIP,L"MenuPopup"},
	{L"SoundCommand",CSetting::TYPE_SOUND,IDS_SOUND_COMMAND,IDS_SOUND_COMMAND_TIP,L"MenuCommand"},
	{L"SoundDrop",CSetting::TYPE_SOUND,IDS_SOUND_DROP,IDS_SOUND_DROP_TIP,L"MoveMenuItem"},

{L"CustomMenu",CSetting::TYPE_GROUP,IDS_CUSTOM_SETTINGS,0,0,0,NULL,&g_CustomMenuPanel},
	{L"MenuItems",CSetting::TYPE_MULTISTRING,0,0,g_DefaultStartMenu},

{L"WindowsMenu",CSetting::TYPE_GROUP,IDS_WSM_SETTINGS},
	{L"CascadeAll",CSetting::TYPE_BOOL,IDS_CASCADE_ALL,IDS_CASCADE_ALL_TIP,0,CSetting::FLAG_BASIC},
	{L"AllProgramsDelay",CSetting::TYPE_INT,IDS_ALL_DELAY,IDS_ALL_DELAY_TIP,-1,0,L"CascadeAll"}, // system hover time
	{L"InitiallySelect",CSetting::TYPE_INT,IDS_ALL_SELECT,IDS_ALL_SELECT_TIP,0,0,L"CascadeAll"},
		{L"SelectSearch",CSetting::TYPE_RADIO,IDS_SELECT_SEARCH,IDS_SELECT_SEARCH_TIP},
		{L"SelectButton",CSetting::TYPE_RADIO,IDS_SELECT_BUTTON,IDS_SELECT_BUTTON_TIP},

{L"AllProgramsSkin",CSetting::TYPE_GROUP,IDS_ALL_SKIN_SETTINGS,0,0,0,NULL,&g_SkinSettingsPanelAll},
	{L"Skin2",CSetting::TYPE_STRING,0,0,L""},
	{L"SkinVariation2",CSetting::TYPE_STRING,0,0,L""},
	{L"SkinOptions2",CSetting::TYPE_STRING,0,0,L""},

{L"Language",CSetting::TYPE_GROUP,IDS_LANGUAGE_SETTINGS_SM,0,0,0,NULL,GetLanguageSettings()},
	{L"Language",CSetting::TYPE_STRING,0,0,L"",CSetting::FLAG_COLD},

{NULL}
};

void UpdateSettings( void )
{
	HDC hdc=::GetDC(NULL);
	int dpi=GetDeviceCaps(hdc,LOGPIXELSY);
	::ReleaseDC(NULL,hdc);
	int iconSize=24;
	if (dpi<=96)
		iconSize=16;
	else if (dpi<=120)
		iconSize=20;
	UpdateSetting(L"SmallIconSize",CComVariant(iconSize),false);
	UpdateSetting(L"LargeIconSize",CComVariant(iconSize*2),false);

	DWORD time;
	SystemParametersInfo(SPI_GETMENUSHOWDELAY,NULL,&time,0);
	UpdateSetting(L"MenuDelay",CComVariant((int)time),false);
	SystemParametersInfo(SPI_GETMOUSEHOVERTIME,NULL,&time,0);
	UpdateSetting(L"AllProgramsDelay",CComVariant((int)time),false);

	int animation=0;
	BOOL animate;
	SystemParametersInfo(SPI_GETMENUANIMATION,NULL,&animate,0);
	if (animate)
	{
		DWORD fade;
		SystemParametersInfo(SPI_GETMENUFADE,NULL,&fade,0);
		animation=fade?1:2;
	}
	UpdateSetting(L"MainMenuAnimation",CComVariant((int)animation),false);
	UpdateSetting(L"SubMenuAnimation",CComVariant((int)animation),false);

	DWORD fade;
	SystemParametersInfo(SPI_GETSELECTIONFADE,NULL,&fade,0);
	UpdateSetting(L"MenuFadeSpeed",CComVariant(fade?400:0),false);

	const wchar_t *skin;
	BOOL comp;
	if (!IsAppThemed())
		skin=L"Classic Skin";
	else if (LOWORD(GetVersion())==0x0006)
		skin=L"Windows Vista Aero";
	else if (SUCCEEDED(DwmIsCompositionEnabled(&comp)) && comp)
		skin=L"Windows 7 Aero";
	else
		skin=L"Windows 7 Basic";
	UpdateSetting(L"Skin1",CComVariant(skin),false);

	UpdateSetting(L"Favorites",CComVariant(0),SHRestricted(REST_NOFAVORITESMENU)!=0);
	UpdateSetting(L"Documents",CComVariant(2),SHRestricted(REST_NORECENTDOCSMENU)!=0);

	DWORD logoff1=SHRestricted(REST_STARTMENULOGOFF);
	DWORD logoff2=SHRestricted(REST_FORCESTARTMENULOGOFF);
	UpdateSetting(L"LogOff",CComVariant((logoff1==2 || logoff2)?1:0),logoff1 || logoff2);

	bool bNoClose=SHRestricted(REST_NOCLOSE)!=0;
	UpdateSetting(L"Shutdown",CComVariant(bNoClose?0:1),bNoClose);
	UpdateSetting(L"RemoteShutdown",CComVariant(0),bNoClose);

	bool bNoUndock=SHRestricted(REST_NOSMEJECTPC)!=0;
	UpdateSetting(L"Undock",CComVariant(bNoUndock?0:1),bNoUndock);

	bool bNoSetFolders=SHRestricted(REST_NOSETFOLDERS)!=0; // hide control panel, printers, network
	bool bNoControlPanel=bNoSetFolders || SHRestricted(REST_NOCONTROLPANEL);
	UpdateSetting(L"ControlPanel",CComVariant(bNoControlPanel?0:1),bNoControlPanel);

	bool bNoNetwork=bNoSetFolders || SHRestricted(REST_NONETWORKCONNECTIONS);
	UpdateSetting(L"Network",CComVariant(bNoNetwork?0:1),bNoNetwork);

	UpdateSetting(L"Printers",CComVariant(bNoSetFolders?0:1),bNoSetFolders);

	bool bNoHelp=SHRestricted(REST_NOSMHELP)!=0;
	UpdateSetting(L"Help",CComVariant(bNoHelp?0:1),bNoHelp);

	bool bNoRun=SHRestricted(REST_NORUN)!=0;
	UpdateSetting(L"Run",CComVariant(bNoRun?0:1),bNoRun);

	bool bNoSearch=SHRestricted(REST_NOFIND)!=0;
	UpdateSetting(L"Search",CComVariant(bNoSearch?0:1),bNoSearch);

	bool bNoDocs=SHRestricted(REST_NOSMMYDOCS)!=0;
	UpdateSetting(L"UserFiles",CComVariant(bNoDocs?0:1),bNoDocs);
	UpdateSetting(L"UserDocuments",CComVariant(bNoDocs?0:1),bNoDocs);
	UpdateSetting(L"UserPictures",CComVariant(bNoDocs?0:1),bNoDocs);
	
	bool bNoEdit=SHRestricted(REST_NOCHANGESTARMENU)!=0;
	UpdateSetting(L"EnableDragDrop",CComVariant(bNoEdit?0:1),bNoEdit);
	UpdateSetting(L"EnableContextMenu",CComVariant(bNoEdit?0:1),bNoEdit);

	UpdateSetting(L"NumericSort",CComVariant(SHRestricted(REST_NOSTRCMPLOGICAL)?0:1),false);

	CRegKey regTitle;
	wchar_t title[256]=L"Windows";
	ULONG size=_countof(title);
	if (regTitle.Open(HKEY_LOCAL_MACHINE,L"Software\\Microsoft\\Windows NT\\CurrentVersion",KEY_READ)==ERROR_SUCCESS)
		regTitle.QueryStringValue(L"ProductName",title,&size);
	UpdateSetting(L"MenuCaption",CComVariant(title),false);

	size=_countof(title);
	if (!GetUserNameEx(NameDisplay,title,&size))
	{
		// GetUserNameEx may fail (for example on Home editions). use the login name
		size=_countof(title);
		GetUserName(title,&size);
	}
	UpdateSetting(L"MenuUsername",CComVariant(title),false);
}

void InitSettings( void )
{
	InitSettings(g_Settings,L"Software\\IvoSoft\\ClassicStartMenu");
}

void ClosingSettings( HWND hWnd, int flags, int command )
{
	EnableHotkeys(HOTKEYS_NORMAL);
	if (command==IDOK)
	{
		if (flags&CSetting::FLAG_COLD)
			MessageBox(hWnd,LoadStringEx(IDS_NEW_SETTINGS),LoadStringEx(IDS_APP_TITLE),MB_OK|MB_ICONWARNING);
	}
}

void EditSettings( bool bModal )
{
#ifndef BUILD_SETUP
	wchar_t path[_MAX_PATH];
	GetModuleFileName(NULL,path,_countof(path));
	if (_wcsicmp(PathFindFileName(path),L"ClassicStartMenu.exe")==0)
		bModal=true;
#endif
	EnableHotkeys(HOTKEYS_SETTINGS);
	wchar_t title[100];
	DWORD ver=GetVersionEx(g_Instance);
	if (ver)
		Sprintf(title,_countof(title),LoadStringEx(IDS_SETTINGS_TITLE_VER),ver>>24,(ver>>16)&0xFF,ver&0xFFFF);
	else
		Sprintf(title,_countof(title),LoadStringEx(IDS_SETTINGS_TITLE));
	EditSettings(title,bModal);
}

void LogHookError( int error )
{
	if (GetSettingInt(L"LogLevel")>0)
	{
		wchar_t fname[_MAX_PATH]=L"%LOCALAPPDATA%\\StartMenuLog.txt";
		DoEnvironmentSubst(fname,_countof(fname));
		FILE *f;
		if (_wfopen_s(&f,fname,L"wb")==0)
		{
			fprintf(f,"Failed to hook Explorer - error=0x%X\r\n",error);
			fclose(f);
		}
	}
}
