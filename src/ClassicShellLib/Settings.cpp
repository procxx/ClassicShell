// Classic Shell (c) 2009-2011, Ivo Beltchev
// The sources for Classic Shell are distributed under the MIT open source license

#include <windows.h>
#include <atlbase.h>
#include <atlwin.h>
#include <atlstr.h>
#include "resource.h"
#include "Settings.h"
#include "SettingsParser.h"
#include "SettingsUIHelper.h"
#include "ResourceHelper.h"
#include "StringUtils.h"
#include "FNVHash.h"
#include <shlobj.h>
#include <Uxtheme.h>
#include <VSStyle.h>
#include <htmlhelp.h>
#include <vector>
#include <algorithm>

#ifdef BUILD_SETUP
#define DOC_PATH L""
#else
#define DOC_PATH L"..\\..\\Docs\\Help\\"
#endif

///////////////////////////////////////////////////////////////////////////////

// Read/Write lock for accessing the settings. Can't be acquired recursively. Only the main UI thread (the one displaying the settings UI)
// can write the settings, and because of that it shouldn't lock when reading the settings. The settings editing code shouldn't use
// GetSettings#### at all to avoid deadlocks
static SRWLOCK g_SettingsLock;

#ifdef _DEBUG
static _declspec(thread) int g_LockState; // 0 - none, 1 - read, 2 - write
static _declspec(thread) bool g_bUIThread; // set to true in the thread that edits the settings
#endif

CSettingsLockRead::CSettingsLockRead( void )
{
#ifdef _DEBUG
	ATLASSERT(g_LockState==0);
	g_LockState=1;
#endif
	AcquireSRWLockShared(&g_SettingsLock);
}

CSettingsLockRead::~CSettingsLockRead( void )
{
#ifdef _DEBUG
	ATLASSERT(g_LockState==1);
	g_LockState=0;
#endif
	ReleaseSRWLockShared(&g_SettingsLock);
}

CSettingsLockWrite::CSettingsLockWrite( void )
{
#ifdef _DEBUG
	ATLASSERT(g_LockState==0);
	g_LockState=2;
#endif
	AcquireSRWLockExclusive(&g_SettingsLock);
}

CSettingsLockWrite::~CSettingsLockWrite( void )
{
#ifdef _DEBUG
	ATLASSERT(g_LockState==2);
	g_LockState=0;
#endif
	ReleaseSRWLockExclusive(&g_SettingsLock);
}

static bool IsVariantTrue( const CComVariant &var )
{
	return (var.vt==VT_I4 && var.intVal==1);
}

bool CSetting::IsEnabled( void ) const
{
	if (flags&CSetting::FLAG_LOCKED_MASK)
		return false;

	if (depend)
	{
		bool checkEnabled=(*depend=='#');
		const wchar_t *name=depend;
		if (checkEnabled)
			name++;

		int len=Strlen(depend);
		int val=0;
		wchar_t operation='~';
		const wchar_t operations[]=L"=~<>";
		for (const wchar_t *c=operations;*c;c++)
		{
			const wchar_t *p=wcschr(depend,*c);
			if (p)
			{
				operation=*c;
				len=(int)(p-depend);
				val=_wtol(p+1);
				break;
			}
		}
		for (const CSetting *pSetting=GetAllSettings();pSetting->name;pSetting++)
		{
			if (_wcsnicmp(pSetting->name,name,len)==0)
			{
				if (checkEnabled && !pSetting->IsEnabled())
					return false;
				if ((pSetting->type==CSetting::TYPE_BOOL || pSetting->type==CSetting::TYPE_INT) && pSetting->value.vt==VT_I4)
				{
					if (operation=='=' && pSetting->value.intVal!=val)
						return false;
					if (operation=='~' && pSetting->value.intVal==val)
						return false;
					if (operation=='<' && pSetting->value.intVal>=val)
						return false;
					if (operation=='>' && pSetting->value.intVal<=val)
						return false;
				}
				break;
			}
		}
	}
	return true;
}

class CSettingsManager
{
public:
	~CSettingsManager( void );
	void Init( CSetting *pSettings, TSettingsComponent component );

	bool GetSettingBool( const wchar_t *name ) const;
	bool GetSettingBool( const wchar_t *name, bool &bDef ) const;
	int GetSettingInt( const wchar_t *name ) const;
	int GetSettingInt( const wchar_t *name, bool &bDef ) const;
	CString GetSettingString( const wchar_t *name ) const;

	void SaveSettings( void );
	void LoadSettings( void );
	bool LoadSettingsXml( const wchar_t *fname );
	bool SaveSettingsXml( const wchar_t *fname );
	void ResetSettings( void );

	CSetting *GetSettings( void ) const { return m_pSettings; }
	HIMAGELIST GetImageList( HWND tree );
	const wchar_t *GetRegPath( void ) const { return m_RegPath; }
	const wchar_t *GetXMLName( void ) const { return m_XMLName; }

private:
	CSetting *m_pSettings;
	HIMAGELIST m_ImageList;
	const wchar_t *m_RegPath;
	const wchar_t *m_CompName;
	const wchar_t *m_XMLName;
};

static CSettingsManager g_SettingsManager;

void CSettingsManager::Init( CSetting *pSettings, TSettingsComponent component )
{
	switch (component)
	{
		case COMPONENT_EXPLORER:
			m_RegPath=L"Software\\IvoSoft\\ClassicExplorer";
			m_CompName=L"Explorer";
			m_XMLName=L"Explorer Settings.xml";
			break;
		case COMPONENT_MENU:
			m_RegPath=L"Software\\IvoSoft\\ClassicStartMenu";
			m_CompName=L"StartMenu";
			m_XMLName=L"Menu Settings.xml";
			break;
		case COMPONENT_IE9:
			m_RegPath=L"Software\\IvoSoft\\ClassicIE9";
			m_CompName=L"IE9";
			m_XMLName=L"IE9 Settings.xml";
			break;
	}
	
	m_pSettings=pSettings;
	InitializeSRWLock(&g_SettingsLock);
	CSettingsLockWrite lock;
	UpdateSettings();
	for (CSetting *pSetting=m_pSettings;pSetting->name;pSetting++)
	{
		if (pSetting->type<0) continue;
#ifdef _DEBUG
		if (pSetting->type==CSetting::TYPE_BOOL)
		{
			ATLASSERT(pSetting->defValue.vt==VT_I4 && (pSetting->defValue.intVal==0 || pSetting->defValue.intVal==1));
		}
		else if (pSetting->type==CSetting::TYPE_INT || pSetting->type==CSetting::TYPE_HOTKEY  || pSetting->type==CSetting::TYPE_HOTKEY_ANY || pSetting->type==CSetting::TYPE_COLOR)
		{
			ATLASSERT(pSetting->defValue.vt==VT_I4);
		}
		else if (pSetting->type>=CSetting::TYPE_STRING)
		{
			ATLASSERT(pSetting->defValue.vt==VT_BSTR);
		}
#endif
		pSetting->value=pSetting->defValue;
		pSetting->flags|=CSetting::FLAG_DEFAULT;
	}
	LoadSettings();
	m_ImageList=NULL;
}

CSettingsManager::~CSettingsManager( void )
{
	if (m_ImageList) ImageList_Destroy(m_ImageList);
	InitializeSRWLock(&g_SettingsLock);
}

bool CSettingsManager::GetSettingBool( const wchar_t *name ) const
{
	for (const CSetting *pSetting=m_pSettings;pSetting->name;pSetting++)
	{
		if (pSetting->type==CSetting::TYPE_BOOL && _wcsicmp(pSetting->name,name)==0)
		{
			CSettingsLockRead lock;
			return IsVariantTrue(pSetting->value);
		}
	}
	ATLASSERT(0);
	return false;
}

bool CSettingsManager::GetSettingBool( const wchar_t *name, bool &bDef ) const
{
	for (const CSetting *pSetting=m_pSettings;pSetting->name;pSetting++)
	{
		if (pSetting->type==CSetting::TYPE_BOOL && _wcsicmp(pSetting->name,name)==0)
		{
			CSettingsLockRead lock;
			bDef=(pSetting->flags&CSetting::FLAG_DEFAULT)!=0;
			return IsVariantTrue(pSetting->value);
		}
	}
	ATLASSERT(0);
	bDef=false;
	return 0;
}

int CSettingsManager::GetSettingInt( const wchar_t *name ) const
{
	for (const CSetting *pSetting=m_pSettings;pSetting->name;pSetting++)
	{
		if ((pSetting->type==CSetting::TYPE_INT || pSetting->type==CSetting::TYPE_HOTKEY || pSetting->type==CSetting::TYPE_HOTKEY_ANY || pSetting->type==CSetting::TYPE_COLOR) && _wcsicmp(pSetting->name,name)==0)
		{
			CSettingsLockRead lock;
			ATLASSERT(pSetting->value.vt==VT_I4);
			return pSetting->value.intVal;
		}
	}
	ATLASSERT(0);
	return 0;
}

int CSettingsManager::GetSettingInt( const wchar_t *name, bool &bDef ) const
{
	for (const CSetting *pSetting=m_pSettings;pSetting->name;pSetting++)
	{
		if (pSetting->type==CSetting::TYPE_INT && _wcsicmp(pSetting->name,name)==0)
		{
			CSettingsLockRead lock;
			ATLASSERT(pSetting->value.vt==VT_I4);
			bDef=(pSetting->flags&CSetting::FLAG_DEFAULT)!=0;
			return pSetting->value.intVal;
		}
	}
	ATLASSERT(0);
	bDef=false;
	return 0;
}

CString CSettingsManager::GetSettingString( const wchar_t *name ) const
{
	for (const CSetting *pSetting=m_pSettings;pSetting->name;pSetting++)
	{
		if (pSetting->type>=CSetting::TYPE_STRING && _wcsicmp(pSetting->name,name)==0)
		{
			CSettingsLockRead lock;
			ATLASSERT(pSetting->value.vt==VT_BSTR);
			return pSetting->value.bstrVal;
		}
	}
	ATLASSERT(0);
	return CString();
}

void CSettingsManager::LoadSettings( void )
{
	ATLASSERT(g_LockState==2);

	// set all to default and unlocked
	for (CSetting *pSetting=m_pSettings;pSetting->name;pSetting++)
	{
		if (pSetting->type==CSetting::TYPE_GROUP || pSetting->type==CSetting::TYPE_RADIO)
			continue;
		pSetting->flags|=CSetting::FLAG_DEFAULT;
		pSetting->flags&=~CSetting::FLAG_LOCKED_REG;
		pSetting->value=pSetting->defValue;
	}

	{
		// load from HKLM, and lock
		CRegKey regSettings;
		if (regSettings.Open(HKEY_LOCAL_MACHINE,m_RegPath,KEY_READ|KEY_WOW64_64KEY)==ERROR_SUCCESS)
		{
			for (CSetting *pSetting=m_pSettings;pSetting->name;pSetting++)
			{
				if (pSetting->type==CSetting::TYPE_GROUP || pSetting->type==CSetting::TYPE_RADIO)
					continue;
				DWORD val;
				if (regSettings.QueryDWORDValue(pSetting->name,val)==ERROR_SUCCESS)
				{
					if (val=='DEFA')
					{
						pSetting->flags|=CSetting::FLAG_LOCKED_REG;
						continue;
					}
					if (pSetting->type==CSetting::TYPE_BOOL)
					{
						pSetting->value=CComVariant(val?1:0);
						pSetting->flags|=CSetting::FLAG_LOCKED_REG;
						pSetting->flags&=~CSetting::FLAG_DEFAULT;
					}
					if (pSetting->type==CSetting::TYPE_INT || pSetting->type==CSetting::TYPE_HOTKEY || pSetting->type==CSetting::TYPE_HOTKEY_ANY || pSetting->type==CSetting::TYPE_COLOR)
					{
						pSetting->value=CComVariant((int)val);
						pSetting->flags|=CSetting::FLAG_LOCKED_REG;
						pSetting->flags&=~CSetting::FLAG_DEFAULT;
					}
				}
				ULONG len;
				if (pSetting->type==CSetting::TYPE_INT && pSetting[1].type==CSetting::TYPE_RADIO)
				{
					if (regSettings.QueryStringValue(pSetting->name,NULL,&len)==ERROR_SUCCESS)
					{
						CString text;
						regSettings.QueryStringValue(pSetting->name,text.GetBuffer(len),&len);
						text.ReleaseBuffer(len);
						val=0;
						for (CSetting *pRadio=pSetting+1;pRadio->type==CSetting::TYPE_RADIO;pRadio++,val++)
						{
							if (_wcsicmp(text,pRadio->name)==0)
							{
								pSetting->value=CComVariant((int)val);
								pSetting->flags|=CSetting::FLAG_LOCKED_REG;
								pSetting->flags&=~CSetting::FLAG_DEFAULT;
								break;
							}
						}
					}
				}
				else if (pSetting->type==CSetting::TYPE_MULTISTRING)
				{
					if (regSettings.QueryMultiStringValue(pSetting->name,NULL,&len)==ERROR_SUCCESS)
					{
						pSetting->value.vt=VT_BSTR;
						pSetting->value.bstrVal=SysAllocStringLen(NULL,len-1);
						regSettings.QueryMultiStringValue(pSetting->name,pSetting->value.bstrVal,&len);
						for (int i=0;i<(int)len-1;i++)
							if (pSetting->value.bstrVal[i]==0)
								pSetting->value.bstrVal[i]='\n';
						pSetting->flags|=CSetting::FLAG_LOCKED_REG;
						pSetting->flags&=~CSetting::FLAG_DEFAULT;
					}
				}
				else if (pSetting->type>=CSetting::TYPE_STRING && regSettings.QueryStringValue(pSetting->name,NULL,&len)==ERROR_SUCCESS)
				{
					pSetting->value.vt=VT_BSTR;
					pSetting->value.bstrVal=SysAllocStringLen(NULL,len-1);
					regSettings.QueryStringValue(pSetting->name,pSetting->value.bstrVal,&len);
					pSetting->flags|=CSetting::FLAG_LOCKED_REG;
					pSetting->flags&=~CSetting::FLAG_DEFAULT;
				}
			}
		}
	}

	{
		// load from HKCU
		CRegKey regSettings;
		if (regSettings.Open(HKEY_CURRENT_USER,m_RegPath)==ERROR_SUCCESS)
		{
			for (CSetting *pSetting=m_pSettings;pSetting->name;pSetting++)
			{
				if (pSetting->type==CSetting::TYPE_GROUP || pSetting->type==CSetting::TYPE_RADIO)
					continue;
				if (pSetting->flags&CSetting::FLAG_LOCKED_REG)
					continue;
				DWORD val;
				if ((pSetting->type==CSetting::TYPE_BOOL || (pSetting->type==CSetting::TYPE_INT && pSetting[1].type!=CSetting::TYPE_RADIO) || pSetting->type==CSetting::TYPE_HOTKEY || pSetting->type==CSetting::TYPE_HOTKEY_ANY || pSetting->type==CSetting::TYPE_COLOR)
					&& regSettings.QueryDWORDValue(pSetting->name,val)==ERROR_SUCCESS)
				{
					if (pSetting->type==CSetting::TYPE_BOOL)
						pSetting->value=CComVariant(val?1:0);
					else
						pSetting->value=CComVariant((int)val);
					pSetting->flags&=~CSetting::FLAG_DEFAULT;
				}
				ULONG len;
				if (pSetting->type==CSetting::TYPE_INT && pSetting[1].type==CSetting::TYPE_RADIO)
				{
					if (regSettings.QueryStringValue(pSetting->name,NULL,&len)==ERROR_SUCCESS)
					{
						CString text;
						regSettings.QueryStringValue(pSetting->name,text.GetBuffer(len),&len);
						text.ReleaseBuffer(len);
						val=0;
						for (CSetting *pRadio=pSetting+1;pRadio->type==CSetting::TYPE_RADIO;pRadio++,val++)
						{
							if (_wcsicmp(text,pRadio->name)==0)
							{
								pSetting->value=CComVariant((int)val);
								pSetting->flags&=~CSetting::FLAG_DEFAULT;
								break;
							}
						}
					}
					else if (regSettings.QueryDWORDValue(pSetting->name,val)==ERROR_SUCCESS)
					{
						pSetting->value=CComVariant((int)val);
						pSetting->flags&=~CSetting::FLAG_DEFAULT;
					}
				}
				else if (pSetting->type==CSetting::TYPE_MULTISTRING)
				{
					if (regSettings.QueryMultiStringValue(pSetting->name,NULL,&len)==ERROR_SUCCESS)
					{
						pSetting->value.vt=VT_BSTR;
						pSetting->value.bstrVal=SysAllocStringLen(NULL,len-1);
						regSettings.QueryMultiStringValue(pSetting->name,pSetting->value.bstrVal,&len);
						for (int i=0;i<(int)len-1;i++)
							if (pSetting->value.bstrVal[i]==0)
								pSetting->value.bstrVal[i]='\n';
						pSetting->flags&=~CSetting::FLAG_DEFAULT;
					}
				}
				else if (pSetting->type>=CSetting::TYPE_STRING && regSettings.QueryStringValue(pSetting->name,NULL,&len)==ERROR_SUCCESS)
				{
					pSetting->value.vt=VT_BSTR;
					pSetting->value.bstrVal=SysAllocStringLen(NULL,len-1);
					regSettings.QueryStringValue(pSetting->name,pSetting->value.bstrVal,&len);
					pSetting->flags&=~CSetting::FLAG_DEFAULT;
				}
			}
		}
	}
}

void CSettingsManager::SaveSettings( void )
{
	// doesn't need to acquire the lock because it can only run from the UI editing code
	ATLASSERT(g_bUIThread);

	// save non-default to HKCU
	CRegKey regSettings;
	if (regSettings.Open(HKEY_CURRENT_USER,m_RegPath)!=ERROR_SUCCESS)
		regSettings.Create(HKEY_CURRENT_USER,m_RegPath);

	for (const CSetting *pSetting=m_pSettings;pSetting->name;pSetting++)
	{
		if (pSetting->type==CSetting::TYPE_GROUP || pSetting->type==CSetting::TYPE_RADIO)
			continue;
		if (pSetting->flags&(CSetting::FLAG_LOCKED_REG|CSetting::FLAG_HIDDEN))
			continue;
		if (pSetting->flags&CSetting::FLAG_DEFAULT)
		{
			regSettings.DeleteValue(pSetting->name);
			continue;
		}
		if (pSetting->type==CSetting::TYPE_BOOL || (pSetting->type==CSetting::TYPE_INT && pSetting[1].type!=CSetting::TYPE_RADIO) || pSetting->type==CSetting::TYPE_HOTKEY || pSetting->type==CSetting::TYPE_HOTKEY_ANY || pSetting->type==CSetting::TYPE_COLOR)
		{
			DWORD val=0;
			if (pSetting->value.vt==VT_I4)
				val=pSetting->value.intVal;
			regSettings.SetDWORDValue(pSetting->name,val);
		}
		if (pSetting->type==CSetting::TYPE_INT && pSetting[1].type==CSetting::TYPE_RADIO)
		{
			DWORD val=0;
			if (pSetting->value.vt==VT_I4)
				val=pSetting->value.intVal;
			for (const CSetting *pRadio=pSetting+1;pRadio->type==CSetting::TYPE_RADIO;pRadio++,val--)
			{
				if (val==0)
				{
					regSettings.SetStringValue(pSetting->name,pRadio->name);
					break;
				}
			}
		}
		if (pSetting->type==CSetting::TYPE_MULTISTRING)
		{
			if (pSetting->value.vt==VT_BSTR)
			{
				int len=Strlen(pSetting->value.bstrVal);
				for (int i=0;i<len;i++)
					if (pSetting->value.bstrVal[i]=='\n')
						pSetting->value.bstrVal[i]=0;
				regSettings.SetMultiStringValue(pSetting->name,pSetting->value.bstrVal);
				for (int i=0;i<len;i++)
					if (pSetting->value.bstrVal[i]==0)
						pSetting->value.bstrVal[i]='\n';
			}
			else
				regSettings.SetMultiStringValue(pSetting->name,L"\0");
		}
		else if (pSetting->type>=CSetting::TYPE_STRING)
		{
			if (pSetting->value.vt==VT_BSTR)
				regSettings.SetStringValue(pSetting->name,pSetting->value.bstrVal);
			else
				regSettings.SetStringValue(pSetting->name,L"");
		}
	}
}

static CComBSTR g_bstrValue(L"value");
static CComBSTR g_bstrTab(L"\n\t");

bool CSettingsManager::LoadSettingsXml( const wchar_t *fname )
{
	CSettingsLockWrite lock;
	CComPtr<IXMLDOMDocument> pDoc;
	if (FAILED(pDoc.CoCreateInstance(L"Msxml2.FreeThreadedDOMDocument"))) return false;
	pDoc->put_async(VARIANT_FALSE);
	VARIANT_BOOL loaded;
	if (pDoc->load(CComVariant(fname),&loaded)!=S_OK || loaded!=VARIANT_TRUE)
		return false;

	CComPtr<IXMLDOMNode> node;
	HRESULT res=pDoc->selectSingleNode(CComBSTR(L"Settings"),&node);
	if (res!=S_OK) return false;

	DWORD ver=0;
	{
		CComVariant value;
		CComQIPtr<IXMLDOMElement> element=node;
		if (!element || element->getAttribute(CComBSTR(L"component"),&value)!=S_OK || value.vt!=VT_BSTR)
			return false;
		if (_wcsicmp(value.bstrVal,m_CompName)!=0)
			return false;

		value.Clear();
		if (element && element->getAttribute(CComBSTR(L"version"),&value)==S_OK && value.vt==VT_BSTR)
		{
			wchar_t token[10];
			const wchar_t *str=GetToken(value.bstrVal,token,_countof(token),L".");
			ver=(_wtol(token)&0xFF)<<24;
			str=GetToken(str,token,_countof(token),L".");
			ver|=(_wtol(token)&0xFF)<<16;
			ver|=_wtol(str)&0xFFFF;
		}
	}

	ResetSettings();
	CComPtr<IXMLDOMNode> child;
	node->get_firstChild(&child);
	while (child)
	{
		CComBSTR name;
		child->get_nodeName(&name);

		for (CSetting *pSetting=g_SettingsManager.GetSettings();pSetting->name;pSetting++)
		{
			if (pSetting->type==CSetting::TYPE_GROUP || pSetting->type==CSetting::TYPE_RADIO)
				continue;
			if (pSetting->type>=0 && _wcsicmp(pSetting->name,name)==0)
			{
				if (pSetting->flags&CSetting::FLAG_LOCKED_REG)
					break;
				if (pSetting->type==CSetting::TYPE_MULTISTRING)
				{
					// load Lines
					CComPtr<IXMLDOMNode> child2;
					child->get_firstChild(&child2);
					std::vector<wchar_t> string;
					while (child2)
					{
						CComBSTR text;
						if (child2->get_text(&text)==S_OK)
						{
							int len=(int)text.Length();
							int pos=(int)string.size();
							string.resize(pos+len+1);
							memcpy(&string[pos],(const wchar_t*)text,len*2);
							string[pos+len]='\n';
						}
						CComPtr<IXMLDOMNode> next;
						child2->get_nextSibling(&next);
						child2=next;
					}
					string.push_back(0);
					pSetting->value=CComVariant(&string[0]);
					pSetting->flags&=~CSetting::FLAG_DEFAULT;
				}
				else
				{
					CComQIPtr<IXMLDOMElement> element=child;
					if (element)
					{
						CComVariant value;
						if (element->getAttribute(g_bstrValue,&value)==S_OK && value.vt==VT_BSTR)
						{
							if (pSetting->type>=CSetting::TYPE_STRING)
							{
								pSetting->value=value;
								pSetting->flags&=~CSetting::FLAG_DEFAULT;
							}
							else if (pSetting->type==CSetting::TYPE_BOOL || (pSetting->type==CSetting::TYPE_INT && pSetting[1].type!=CSetting::TYPE_RADIO) || pSetting->type==CSetting::TYPE_HOTKEY || pSetting->type==CSetting::TYPE_HOTKEY_ANY || pSetting->type==CSetting::TYPE_COLOR)
							{
								int val=_wtol(value.bstrVal);
								if (pSetting->type==CSetting::TYPE_BOOL)
									pSetting->value=CComVariant(val?1:0);
								else
									pSetting->value=CComVariant(val);
								pSetting->flags&=~CSetting::FLAG_DEFAULT;
							}
							else if (pSetting->type==CSetting::TYPE_INT && pSetting[1].type==CSetting::TYPE_RADIO)
							{
								int val=0;
								for (CSetting *pRadio=pSetting+1;pRadio->type==CSetting::TYPE_RADIO;pRadio++,val++)
								{
									if (_wcsicmp(pRadio->name,value.bstrVal)==0)
									{
										pSetting->value=CComVariant(val);
										pSetting->flags&=~CSetting::FLAG_DEFAULT;
										break;
									}
								}
							}
						}
					}
				}
			}
		}

		CComPtr<IXMLDOMNode> next;
		if (child->get_nextSibling(&next)!=S_OK)
			break;
		child=next;
	}
	return true;
}

static void SaveSettingValue( IXMLDOMDocument *pDoc, IXMLDOMNode *pParent, const wchar_t *name, const CComVariant &value )
{
	CComPtr<IXMLDOMElement> setting;
	pDoc->createElement(CComBSTR(name),&setting);
	setting->setAttribute(g_bstrValue,value);
	CComPtr<IXMLDOMText> text;
	CComPtr<IXMLDOMNode> nu;
	pDoc->createTextNode(g_bstrTab,&text);
	pParent->appendChild(text,&nu);
	nu=NULL;
	pParent->appendChild(setting,&nu);
}

bool CSettingsManager::SaveSettingsXml( const wchar_t *fname )
{
	// doesn't need to acquire the lock because it can only run from the UI editing code
	ATLASSERT(g_bUIThread);

	CComPtr<IXMLDOMDocument> pDoc;
	HRESULT res=pDoc.CoCreateInstance(L"Msxml2.FreeThreadedDOMDocument");
	if (FAILED(res))
		return false;

	CComPtr<IXMLDOMElement> pRoot;
	pDoc->createElement(CComBSTR(L"Settings"),&pRoot);
	CComPtr<IXMLDOMProcessingInstruction> pi;
	if (SUCCEEDED(pDoc->createProcessingInstruction(CComBSTR(L"xml"),CComBSTR(L"version=\"1.0\""),&pi)))
	{
		CComPtr<IXMLDOMNode> nu;
		pDoc->appendChild(pi,&nu);
	}

	{
		CComPtr<IXMLDOMNode> nu;
		pDoc->appendChild(pRoot,&nu);
	}

	pRoot->setAttribute(CComBSTR(L"component"),CComVariant(m_CompName));

	wchar_t version[100];
	DWORD ver=GetVersionEx(_AtlBaseModule.GetResourceInstance());
	Sprintf(version,_countof(version),L"%d.%d.%d",ver>>24,(ver>>16)&0xFF,ver&0xFFFF);
	pRoot->setAttribute(CComBSTR(L"version"),CComVariant(version));

	for (const CSetting *pSetting=m_pSettings;pSetting->name;pSetting++)
	{
		if (pSetting->type==CSetting::TYPE_GROUP || pSetting->type==CSetting::TYPE_RADIO)
			continue;
		if (pSetting->flags&(CSetting::FLAG_LOCKED_REG|CSetting::FLAG_HIDDEN|CSetting::FLAG_DEFAULT))
			continue;
		if (pSetting->type==CSetting::TYPE_MULTISTRING)
		{
			CComPtr<IXMLDOMElement> setting;
			pDoc->createElement(CComBSTR(pSetting->name),&setting);
			CComPtr<IXMLDOMText> text;
			CComPtr<IXMLDOMNode> nu;
			pDoc->createTextNode(g_bstrTab,&text);
			pRoot->appendChild(text,&nu);
			nu=NULL;
			pRoot->appendChild(setting,&nu);
			CComBSTR tabs(L"\n\t\t");
			CComBSTR name(L"Line");
			if (pSetting->value.vt==VT_BSTR)
			{
				for (const wchar_t *str=pSetting->value.bstrVal;*str;)
				{
					int len;
					const wchar_t *end=wcschr(str,'\n');
					if (end)
						len=(int)(end-str);
					else
						len=Strlen(str);
					CComPtr<IXMLDOMElement> line;
					pDoc->createElement(name,&line);
					line->put_text(CComBSTR(len,str));
					nu=NULL;
					text=NULL;
					pDoc->createTextNode(tabs,&text);
					setting->appendChild(text,&nu);
					nu=NULL;
					setting->appendChild(line,&nu);
					if (!end) break;
					str=end+1;
				}
			}
			nu=NULL;
			text=NULL;
			pDoc->createTextNode(g_bstrTab,&text);
			setting->appendChild(text,&nu);
			continue;
		}
		else if (pSetting->type==CSetting::TYPE_BOOL || (pSetting->type==CSetting::TYPE_INT && pSetting[1].type!=CSetting::TYPE_RADIO) || pSetting->type>=CSetting::TYPE_HOTKEY || pSetting->type>=CSetting::TYPE_HOTKEY_ANY || pSetting->type>=CSetting::TYPE_STRING)
		{
			SaveSettingValue(pDoc,pRoot,pSetting->name,pSetting->value);
		}
		else if (pSetting->type==CSetting::TYPE_INT && pSetting[1].type==CSetting::TYPE_RADIO)
		{
			DWORD val=0;
			if (pSetting->value.vt==VT_I4)
				val=pSetting->value.intVal;
			for (const CSetting *pRadio=pSetting+1;pRadio->type==CSetting::TYPE_RADIO;pRadio++,val--)
			{
				if (val==0)
				{
					SaveSettingValue(pDoc,pRoot,pSetting->name,CComVariant(pRadio->name));
					break;
				}
			}
		}
	}
	CComPtr<IXMLDOMText> text;
	CComPtr<IXMLDOMNode> nu;
	pDoc->createTextNode(CComBSTR(L"\n"),&text);
	pRoot->appendChild(text,&nu);
	return SUCCEEDED(pDoc->save(CComVariant(fname)));
}

void CSettingsManager::ResetSettings( void )
{
	ATLASSERT(g_LockState==2); // must be locked for writing
	for (CSetting *pSetting=m_pSettings;pSetting->name;pSetting++)
	{
		if (pSetting->type==CSetting::TYPE_GROUP || pSetting->type==CSetting::TYPE_RADIO)
			continue;
		if (pSetting->flags&CSetting::FLAG_LOCKED_REG)
			continue;
		pSetting->flags|=CSetting::FLAG_DEFAULT;
		pSetting->value=pSetting->defValue;
	}
}

HIMAGELIST CSettingsManager::GetImageList( HWND tree )
{
	if (m_ImageList) return m_ImageList;
	int size=TreeView_GetItemHeight(tree);
	if (size>16 && size<20) size=16; // avoid weird sizes that can distort the icons
	m_ImageList=ImageList_Create(size,size,ILC_COLOR32|ILC_MASK|((GetWindowLong(tree,GWL_EXSTYLE)&WS_EX_LAYOUTRTL)?ILC_MIRROR:0),0,23);
	BITMAPINFO dib={sizeof(dib)};
	dib.bmiHeader.biWidth=size;
	dib.bmiHeader.biHeight=-size;
	dib.bmiHeader.biPlanes=1;
	dib.bmiHeader.biBitCount=32;
	dib.bmiHeader.biCompression=BI_RGB;
	HDC hdc=CreateCompatibleDC(NULL);
	HDC hdcMask=CreateCompatibleDC(NULL);
	HBITMAP bmp=CreateDIBSection(hdc,&dib,DIB_RGB_COLORS,NULL,NULL,0);
	HBITMAP bmpMask=CreateDIBSection(hdcMask,&dib,DIB_RGB_COLORS,NULL,NULL,0);

	HTHEME theme=OpenThemeData(tree,L"button");

	for (int i=0;i<13;i++)
	{
		HBITMAP bmp0=(HBITMAP)SelectObject(hdc,bmp);
		HBITMAP bmp1=(HBITMAP)SelectObject(hdcMask,bmpMask);
		RECT rc={0,0,size,size};
		FillRect(hdc,&rc,(HBRUSH)(COLOR_WINDOW+1));
		FillRect(hdcMask,&rc,(HBRUSH)GetStockObject(BLACK_BRUSH));
		if (i==1)
		{
			HICON icon=(HICON)LoadImage(_AtlBaseModule.GetResourceInstance(),MAKEINTRESOURCE(IDI_ICONLOCK),IMAGE_ICON,size,size,LR_DEFAULTCOLOR);
			DrawIconEx(hdc,0,0,icon,size,size,0,NULL,DI_NORMAL);
			DrawIconEx(hdcMask,0,0,icon,size,size,0,NULL,DI_MASK);
			DestroyIcon(icon);
		}
		else if (i==2 || i==3)
		{
			HMODULE hShell32=GetModuleHandle(L"shell32.dll");
			if (hShell32)
			{
				HICON icon=(HICON)LoadImage(hShell32,MAKEINTRESOURCE(16775),IMAGE_ICON,size,size,LR_DEFAULTCOLOR);
				DrawIconEx(hdc,0,0,icon,size,size,0,NULL,DI_NORMAL);
				DestroyIcon(icon);
			}
		}
		else if (i==12)
		{
			HICON icon=(HICON)LoadImage(_AtlBaseModule.GetResourceInstance(),MAKEINTRESOURCE(IDI_ICONWARNING),IMAGE_ICON,size,size,LR_DEFAULTCOLOR);
			DrawIconEx(hdc,0,0,icon,size,size,0,NULL,DI_NORMAL);
			DrawIconEx(hdcMask,0,0,icon,size,size,0,NULL,DI_MASK);
			DestroyIcon(icon);
		}
		else if (i>3)
		{
			InflateRect(&rc,-1,-1);
			if (theme)
			{
				if ((i-4)&4)
				{
					int state=(i-4)&3;
					if (state==0) state=RBS_UNCHECKEDNORMAL;
					else if (state==1) state=RBS_UNCHECKEDDISABLED;
					else if (state==2) state=RBS_CHECKEDNORMAL;
					else state=RBS_CHECKEDDISABLED;
					DrawThemeBackground(theme,hdc,BP_RADIOBUTTON,state,&rc,NULL);
				}
				else
				{
					int state=(i-4)&3;
					if (state==0) state=CBS_UNCHECKEDNORMAL;
					else if (state==1) state=CBS_UNCHECKEDDISABLED;
					else if (state==2) state=CBS_CHECKEDNORMAL;
					else state=CBS_CHECKEDDISABLED;
					DrawThemeBackground(theme,hdc,BP_CHECKBOX,state,&rc,NULL);
				}
			}
			else
			{
				UINT state=DFCS_BUTTONCHECK|DFCS_FLAT;
				if ((i-4)&1) state|=DFCS_INACTIVE;
				if ((i-4)&2) state|=DFCS_CHECKED;
				if ((i-4)&4) state|=DFCS_BUTTONRADIO;
				DrawFrameControl(hdc,&rc,DFC_BUTTON,state);
			}
		}
		SelectObject(hdc,bmp0);
		SelectObject(hdcMask,bmp1);
		ImageList_Add(m_ImageList,bmp,bmpMask);
	}

	// create color images
	{
		HBITMAP bmp0=(HBITMAP)SelectObject(hdc,bmp);
		HBITMAP bmp1=(HBITMAP)SelectObject(hdcMask,bmpMask);
		RECT rc={0,0,size,size};
		FillRect(hdc,&rc,(HBRUSH)GetStockObject(BLACK_BRUSH));
		FillRect(hdcMask,&rc,(HBRUSH)GetStockObject(BLACK_BRUSH));
		SelectObject(hdc,bmp0);
		SelectObject(hdcMask,bmp1);

		for (int i=0;i<10;i++)
			ImageList_Add(m_ImageList,bmp,bmpMask);
	}

	DeleteObject(bmp);
	DeleteObject(bmpMask);
	DeleteDC(hdc);
	DeleteDC(hdcMask);

	if (theme) CloseThemeData(theme);
	ImageList_SetOverlayImage(m_ImageList,1,1);
	ImageList_SetOverlayImage(m_ImageList,12,2);
	return m_ImageList;
}

///////////////////////////////////////////////////////////////////////////////

class CSettingsDlg: public CResizeableDlg<CSettingsDlg>
{
public:
	void Init( CSetting *pSettings );

	BEGIN_MSG_MAP( CSettingsDlg )
		MESSAGE_HANDLER( WM_INITDIALOG, OnInitDialog )
		MESSAGE_HANDLER( WM_DESTROY, OnDestroy )
		MESSAGE_HANDLER( WM_SIZE, OnSize )
		MESSAGE_HANDLER( WM_GETMINMAXINFO, OnGetMinMaxInfo )
		MESSAGE_HANDLER( WM_KEYDOWN, OnKeyDown )
		MESSAGE_HANDLER( WM_SYSCOMMAND, OnSysCommand )
		COMMAND_HANDLER( IDOK, BN_CLICKED, OnOK )
		COMMAND_HANDLER( IDCANCEL, BN_CLICKED, OnCancel )
		COMMAND_HANDLER( IDC_BUTTONBACKUP, BN_CLICKED, OnBackup )
		COMMAND_HANDLER( IDC_RADIOBASIC, BN_CLICKED, OnBasic )
		COMMAND_HANDLER( IDC_RADIOALL, BN_CLICKED, OnBasic )
		NOTIFY_HANDLER( IDC_TABSETTINGS, TCN_SELCHANGING, OnSelChanging )
		NOTIFY_HANDLER( IDC_TABSETTINGS, TCN_SELCHANGE, OnSelChange )
		NOTIFY_HANDLER( IDC_BUTTONBACKUP, BCN_DROPDOWN, OnDropDown )
		NOTIFY_HANDLER( IDC_LINKHELP, NM_CLICK, OnHelp )
		NOTIFY_HANDLER( IDC_LINKHELP, NM_RETURN, OnHelp )
	END_MSG_MAP()

	BEGIN_RESIZE_MAP
		RESIZE_CONTROL(IDC_TABSETTINGS,MOVE_SIZE_X|MOVE_SIZE_Y)
		RESIZE_CONTROL(IDC_LINKHELP,MOVE_MOVE_Y)
		RESIZE_CONTROL(IDC_RADIOBASIC,MOVE_MOVE_Y)
		RESIZE_CONTROL(IDC_RADIOALL,MOVE_MOVE_Y)
		RESIZE_CONTROL(IDC_BUTTONBACKUP,MOVE_MOVE_X|MOVE_MOVE_Y)
		RESIZE_CONTROL(IDOK,MOVE_MOVE_X|MOVE_MOVE_Y)
		RESIZE_CONTROL(IDCANCEL,MOVE_MOVE_X|MOVE_MOVE_Y)
	END_RESIZE_MAP

	bool GetOnTop( void ) const { return m_bOnTop; }
	void ShowHelp( void ) const;

protected:
	// Handler prototypes:
	//  LRESULT MessageHandler(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
	//  LRESULT CommandHandler(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled);
	//  LRESULT NotifyHandler(int idCtrl, LPNMHDR pnmh, BOOL& bHandled);
	LRESULT OnInitDialog( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled );
	LRESULT OnDestroy( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled );
	LRESULT OnSize( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled );
	LRESULT OnKeyDown( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled );
	LRESULT OnSysCommand( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled );
	LRESULT OnOK( WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled );
	LRESULT OnCancel( WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled );
	LRESULT OnBackup( WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled );
	LRESULT OnBasic( WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled );
	LRESULT OnSelChanging( int idCtrl, LPNMHDR pnmh, BOOL& bHandled );
	LRESULT OnSelChange( int idCtrl, LPNMHDR pnmh, BOOL& bHandled );
	LRESULT OnDropDown( int idCtrl, LPNMHDR pnmh, BOOL& bHandled );
	LRESULT OnHelp( int idCtrl, LPNMHDR pnmh, BOOL& bHandled );

private:
	CSetting *m_pSettings;
	CWindow m_Tabs;
	int m_Index;
	HWND m_Panel;
	bool m_bBasic;
	bool m_bOnTop;

	void AddTabs( int name );
	void SetCurTab( int index, bool bReset );
	bool IsTabValid( void );
	void StorePlacement( void );

	struct Placement
	{
		RECT rc;
		unsigned int tab;
		bool basic;
		bool top;
	};
};

void CSettingsDlg::Init( CSetting *pSettings )
{
	m_pSettings=pSettings;
}

LRESULT CSettingsDlg::OnInitDialog( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled )
{
#ifdef _DEBUG
	g_bUIThread=true;
#endif
	CResizeableDlg<CSettingsDlg>::InitResize(MOVE_MODAL);
	HMENU menu=GetSystemMenu(FALSE);
	bool bAdded=false;
	int n=GetMenuItemCount(menu);
	for (int i=0;i<n;i++)
	{
		int id=GetMenuItemID(menu,i);
		if (id==SC_MAXIMIZE)
		{
			InsertMenu(menu,i+1,MF_BYPOSITION|MF_STRING,1,LoadStringEx(IDS_ALWAYS_ON_TOP));
			bAdded=true;
			break;
		}
	}
	if (!bAdded)
		InsertMenu(menu,SC_CLOSE,MF_BYCOMMAND|MF_STRING,1,LoadStringEx(IDS_ALWAYS_ON_TOP));

	Placement pos;
	CRegKey regSettings;
	if (regSettings.Open(HKEY_CURRENT_USER,GetSettingsRegPath())==ERROR_SUCCESS)
	{
		ULONG size=sizeof(pos);
		if (regSettings.QueryBinaryValue(L"CSettingsDlg",&pos,&size)!=ERROR_SUCCESS || size!=sizeof(pos))
		{
			memset(&pos,0,sizeof(pos));
			pos.basic=true;
		}
		regSettings.Close();
	}
	else
	{
		memset(&pos,0,sizeof(pos));
		pos.basic=true;
	}

	m_bOnTop=pos.top;
	if (m_bOnTop)
	{
		CheckMenuItem(menu,1,MF_BYCOMMAND|MF_CHECKED);
	}

	HINSTANCE hInstance=_AtlBaseModule.GetResourceInstance();

	HICON icon=(HICON)LoadImage(hInstance,MAKEINTRESOURCE(1),IMAGE_ICON,GetSystemMetrics(SM_CXICON),GetSystemMetrics(SM_CYICON),LR_DEFAULTCOLOR);
	SendMessage(WM_SETICON,ICON_BIG,(LPARAM)icon);
	icon=(HICON)LoadImage(hInstance,MAKEINTRESOURCE(1),IMAGE_ICON,GetSystemMetrics(SM_CXSMICON),GetSystemMetrics(SM_CYSMICON),LR_DEFAULTCOLOR);
	SendMessage(WM_SETICON,ICON_SMALL,(LPARAM)icon);

	CWindow backup=GetDlgItem(IDC_BUTTONBACKUP);
	backup.SetWindowLong(GWL_STYLE,backup.GetWindowLong(GWL_STYLE)|BS_SPLITBUTTON);
	BUTTON_SPLITINFO info={BCSIF_STYLE,NULL,BCSS_NOSPLIT};
	backup.SendMessage(BCM_SETSPLITINFO,0,(LPARAM)&info);

	CWindow parent=GetParent();
	if (parent)
	{
		RECT rc1,rc2;
		GetWindowRect(&rc1);
		parent.GetWindowRect(&rc2);
		OffsetRect(&rc1,(rc2.left+rc2.right)/2-(rc1.left+rc1.right)/2,(rc2.top+rc2.bottom)/2-(rc1.top+rc1.bottom)/2);
		if (rc1.top<rc2.top) OffsetRect(&rc1,0,rc2.top-rc1.top);
		SetWindowPos(NULL,rc1.left,rc1.top,rc1.right-rc1.left,rc1.bottom-rc1.top,SWP_NOZORDER);
		SendMessage(DM_REPOSITION,0,0);
	}

	m_bBasic=pos.basic;
	CheckDlgButton(IDC_RADIOBASIC,m_bBasic?BST_CHECKED:BST_UNCHECKED);
	CheckDlgButton(IDC_RADIOALL,m_bBasic?BST_UNCHECKED:BST_CHECKED);

	SIZE maxSize={0,0};
	m_Tabs=GetDlgItem(IDC_TABSETTINGS);
	m_Panel=NULL;
	int idx=0;
	for (const CSetting *pSetting=m_pSettings;pSetting->name;pSetting++)
	{
		if (pSetting->type!=CSetting::TYPE_GROUP)
			continue;
		ISettingsPanel *pPanel=pSetting->pPanel;
		if (!pPanel) pPanel=GetDefaultSettings();
		HWND panel=pPanel->Create(m_hWnd);
		RECT rc;
		::GetWindowRect(panel,&rc);
		if (maxSize.cx<rc.right-rc.left)
			maxSize.cx=rc.right-rc.left;
		if (maxSize.cy<rc.bottom-rc.top)
			maxSize.cy=rc.bottom-rc.top;
	}

	RECT rc;
	m_Tabs.GetWindowRect(&rc);
	TabCtrl_AdjustRect(m_Tabs,FALSE,&rc);
	int dx=maxSize.cx-(rc.right-rc.left);
	int dy=maxSize.cy-(rc.bottom-rc.top);
	if (dx<0) dx=0;
	if (dy<0) dy=0;
	if (dx || dy)
	{
		GetWindowRect(&rc);
		rc.left-=dx/2;
		rc.right+=dx-dx/2;
		rc.top-=dy/2;
		rc.bottom+=dy-dy/2;
		SetWindowPos(NULL,&rc,SWP_NOZORDER);
		CResizeableDlg<CSettingsDlg>::InitResize(MOVE_MODAL|MOVE_REINITIALIZE);
	}

	{
		CSettingsLockWrite lock;
		for (CSetting *pSetting=m_pSettings;pSetting->name;pSetting++)
		{
			pSetting->tempValue=pSetting->value;
			pSetting->tempFlags=pSetting->flags;
		}
	}

	AddTabs(pos.tab);
	if (pos.tab)
		SetStoreRect(pos.rc);

	return TRUE;
}

LRESULT CSettingsDlg::OnDestroy( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled )
{
	bHandled=FALSE;
#ifdef _DEBUG
	g_bUIThread=false;
#endif
	return 0;
}

void CSettingsDlg::AddTabs( int name )
{
	TabCtrl_DeleteAllItems(m_Tabs);
	int idx=0;
	for (const CSetting *pSetting=m_pSettings;pSetting->name;pSetting++)
	{
		if (pSetting->type!=CSetting::TYPE_GROUP || (pSetting->flags&CSetting::FLAG_HIDDEN))
			continue;
		if (!m_bBasic && pSetting->nameID==IDS_BASIC_SETTINGS)
			continue;
		if (m_bBasic && pSetting->nameID!=IDS_BASIC_SETTINGS && !(pSetting->flags&CSetting::FLAG_BASIC))
			continue;
		CString str=LoadStringEx(pSetting->nameID);
		TCITEM tab={TCIF_PARAM|TCIF_TEXT,0,0,(LPWSTR)(LPCWSTR)str,0,0,(LPARAM)pSetting};
		int i=TabCtrl_InsertItem(m_Tabs,1000,&tab);
		if (pSetting->nameID==name)
			idx=i;
	}
	m_Index=-1;
	TabCtrl_SetCurSel(m_Tabs,idx);
	m_Tabs.InvalidateRect(NULL);
	SetCurTab(idx,false);
}

LRESULT CSettingsDlg::OnSize( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled )
{
	CResizeableDlg<CSettingsDlg>::OnSize();
	RECT rc;
	m_Tabs.GetWindowRect(&rc);
	::MapWindowPoints(NULL,m_hWnd,(POINT*)&rc,2);
	TabCtrl_AdjustRect(m_Tabs,FALSE,&rc);
	if (m_Panel) ::SetWindowPos(m_Panel,HWND_TOP,rc.left,rc.top,rc.right-rc.left,rc.bottom-rc.top,0);
	return 0;
}

LRESULT CSettingsDlg::OnKeyDown( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled )
{
	if (wParam==VK_TAB && GetKeyState(VK_CONTROL)<0)
	{
		int sel=TabCtrl_GetCurSel(m_Tabs);
		if (GetKeyState(VK_SHIFT)<0)
		{
			if (sel>0)
			{
				TabCtrl_SetCurSel(m_Tabs,sel-1);
				SetCurTab(sel-1,false);
			}
		}
		else
		{
			if (sel<TabCtrl_GetItemCount(m_Tabs)-1)
			{
				TabCtrl_SetCurSel(m_Tabs,sel+1);
				SetCurTab(sel+1,false);
			}
		}
	}
	return 0;
}

LRESULT CSettingsDlg::OnSysCommand( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled )
{
	if (wParam==1)
	{
		HMENU menu=GetSystemMenu(FALSE);
		m_bOnTop=!m_bOnTop;
		CheckMenuItem(menu,1,MF_BYCOMMAND|(m_bOnTop?MF_CHECKED:MF_UNCHECKED));
		SetWindowPos(m_bOnTop?HWND_TOPMOST:HWND_NOTOPMOST,0,0,0,0,SWP_NOSIZE|SWP_NOMOVE);
		return 0;
	}
	bHandled=FALSE;
	return 0;
}

void CSettingsDlg::StorePlacement( void )
{
	Placement pos;
	GetStoreRect(pos.rc);
	int idx=TabCtrl_GetCurSel(m_Tabs);
	if (idx<0) return;
	TCITEM tab={TCIF_PARAM};
	TabCtrl_GetItem(m_Tabs,idx,&tab);
	CSetting *pGroup=(CSetting*)tab.lParam;
	pos.tab=pGroup->nameID;
	pos.basic=m_bBasic;
	pos.top=m_bOnTop;

	CRegKey regSettings;
	if (regSettings.Open(HKEY_CURRENT_USER,GetSettingsRegPath())!=ERROR_SUCCESS)
		regSettings.Create(HKEY_CURRENT_USER,GetSettingsRegPath());
	regSettings.SetBinaryValue(L"CSettingsDlg",&pos,sizeof(pos));
}

LRESULT CSettingsDlg::OnOK( WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled )
{
	if (IsTabValid())
	{
		int flags=0;
		for (const CSetting *pSetting=m_pSettings;pSetting->name;pSetting++)
		{
			if (pSetting->value!=pSetting->tempValue)
				flags|=pSetting->flags&(CSetting::FLAG_WARM|CSetting::FLAG_COLD);
		}
		g_SettingsManager.SaveSettings();
		ClosingSettings(m_hWnd,flags,IDOK);
		StorePlacement();
		DestroyWindow();
	}
	return TRUE;
}

LRESULT CSettingsDlg::OnCancel( WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled )
{
	StorePlacement();
	DestroyWindow();
	// restore all settings
	{
		CSettingsLockWrite lock;
		for (CSetting *pSetting=m_pSettings;pSetting->name;pSetting++)
		{
			pSetting->value=pSetting->tempValue;
			pSetting->flags=pSetting->tempFlags;
		}
	}
	ClosingSettings(m_hWnd,0,IDCANCEL);
	return TRUE;
}

LRESULT CSettingsDlg::OnBackup( WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled )
{
	HMENU menu=CreatePopupMenu();
	AppendMenu(menu,MF_STRING,1,LoadStringEx(IDS_BACKUP_SAVE));
	AppendMenu(menu,MF_STRING,2,LoadStringEx(IDS_BACKUP_LOAD));
	AppendMenu(menu,MF_STRING,3,LoadStringEx(IDS_BACKUP_RESET));
	TPMPARAMS params={sizeof(params)};
	GetDlgItem(IDC_BUTTONBACKUP).GetWindowRect(&params.rcExclude);
	if (GetWindowLong(GWL_EXSTYLE)&WS_EX_LAYOUTRTL)
	{
		int q=params.rcExclude.left;
		params.rcExclude.left=params.rcExclude.right;
		params.rcExclude.right=q;
	}
	int res=TrackPopupMenuEx(menu,TPM_RETURNCMD|TPM_VERTICAL,params.rcExclude.left,params.rcExclude.bottom,m_hWnd,&params);
	DestroyMenu(menu);

	// remove the next mouse click if it is on the Backup button
	MSG msg;
	if (PeekMessage(&msg,NULL,WM_LBUTTONDOWN,WM_LBUTTONDBLCLK,PM_NOREMOVE) && PtInRect(&params.rcExclude,msg.pt))
		PeekMessage(&msg,NULL,WM_LBUTTONDOWN,WM_LBUTTONDBLCLK,PM_REMOVE);

	if (res==1)
	{
		// save
		wchar_t path[_MAX_PATH];
		Strcpy(path,_countof(path),g_SettingsManager.GetXMLName());
		OPENFILENAME ofn={sizeof(ofn)};
		ofn.hwndOwner=m_hWnd;
		wchar_t filters[256];
		Strcpy(filters,_countof(filters),LoadStringEx(IDS_XML_FILTERS));
		for (wchar_t *c=filters;*c;c++)
			if (*c=='|') *c=0;
		ofn.lpstrFilter=filters;
		ofn.nFilterIndex=1;
		ofn.lpstrFile=path;
		ofn.nMaxFile=_MAX_PATH;
		CString title=LoadStringEx(IDS_XML_TITLE_SAVE);
		ofn.lpstrTitle=title;
		ofn.lpstrDefExt=L".xml";
		ofn.Flags=OFN_DONTADDTORECENT|OFN_ENABLESIZING|OFN_EXPLORER|OFN_PATHMUSTEXIST|OFN_OVERWRITEPROMPT|OFN_HIDEREADONLY|OFN_NOCHANGEDIR;
		if (GetSaveFileName(&ofn))
		{
			if (!g_SettingsManager.SaveSettingsXml(path))
			{
				wchar_t text[1024];
				Sprintf(text,_countof(text),LoadStringEx(IDS_ERROR_SAVING_XML),path);
				::MessageBox(m_hWnd,text,LoadStringEx(IDS_ERROR_TITLE),MB_OK|MB_ICONERROR);
			}
		}
	}
	if (res==2)
	{
		// load
		wchar_t path[_MAX_PATH];
		path[0]=0;
		OPENFILENAME ofn={sizeof(ofn)};
		ofn.hwndOwner=m_hWnd;
		wchar_t filters[256];
		Strcpy(filters,_countof(filters),LoadStringEx(IDS_XML_FILTERS));
		for (wchar_t *c=filters;*c;c++)
			if (*c=='|') *c=0;
		ofn.lpstrFilter=filters;
		ofn.nFilterIndex=1;
		ofn.lpstrFile=path;
		ofn.nMaxFile=_MAX_PATH;
		CString title=LoadStringEx(IDS_XML_TITLE_LOAD);
		ofn.lpstrTitle=title;
		ofn.Flags=OFN_DONTADDTORECENT|OFN_ENABLESIZING|OFN_EXPLORER|OFN_FILEMUSTEXIST|OFN_HIDEREADONLY|OFN_NOCHANGEDIR;
		if (GetOpenFileName(&ofn))
		{
			if (!g_SettingsManager.LoadSettingsXml(path))
			{
				wchar_t text[1024];
				Sprintf(text,_countof(text),LoadStringEx(IDS_ERROR_LOADING_XML),path);
				::MessageBox(m_hWnd,text,LoadStringEx(IDS_ERROR_TITLE),MB_OK|MB_ICONERROR);
			}
			SetCurTab(m_Index,true);
		}
	}
	if (res==3)
	{
		// reset
		if (::MessageBox(m_hWnd,LoadStringEx(IDS_RESET_CONFIRM),LoadStringEx(IDS_RESET_TITLE),MB_YESNO|MB_ICONWARNING)==IDYES)
		{
			{
				CSettingsLockWrite lock;
				g_SettingsManager.ResetSettings();
			}
			SetCurTab(m_Index,true);
		}
	}
	return TRUE;
}

LRESULT CSettingsDlg::OnBasic( WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled )
{
	bool bBasic=IsDlgButtonChecked(IDC_RADIOBASIC)==BST_CHECKED;
	if (m_bBasic!=bBasic)
	{
		m_bBasic=bBasic;
		AddTabs(-1);
	}
	return 0;
}

void CSettingsDlg::SetCurTab( int index, bool bReset )
{
	if (m_Index==index && !bReset) return;
	m_Index=index;
	TCITEM tab={TCIF_PARAM};
	TabCtrl_GetItem(m_Tabs,index,&tab);
	CSetting *pGroup=(CSetting*)tab.lParam;
	ISettingsPanel *pPanel=pGroup->pPanel;
	if (!pPanel) pPanel=GetDefaultSettings();
	RECT rc;
	m_Tabs.GetWindowRect(&rc);
	::MapWindowPoints(NULL,m_hWnd,(POINT*)&rc,2);
	TabCtrl_AdjustRect(m_Tabs,FALSE,&rc);
	HWND hwnd=pPanel->Activate(pGroup,rc,bReset);
	if (hwnd!=m_Panel)
	{
		if (m_Panel) ::ShowWindow(m_Panel,SW_HIDE);
		m_Panel=hwnd;
		::SetFocus(m_Panel);
	}
}

LRESULT CSettingsDlg::OnSelChanging( int idCtrl, LPNMHDR pnmh, BOOL& bHandled )
{
	return !IsTabValid();
}

LRESULT CSettingsDlg::OnSelChange( int idCtrl, LPNMHDR pnmh, BOOL& bHandled )
{
	SetCurTab(TabCtrl_GetCurSel(m_Tabs),false);
	return 0;
}

LRESULT CSettingsDlg::OnDropDown( int idCtrl, LPNMHDR pnmh, BOOL& bHandled )
{
	SendMessage(WM_COMMAND,IDC_BUTTONBACKUP);
	return 0;
}

LRESULT CSettingsDlg::OnHelp( int idCtrl, LPNMHDR pnmh, BOOL& bHandled )
{
	ShowHelp();
	return 0;
}

void CSettingsDlg::ShowHelp( void ) const
{
	wchar_t path[_MAX_PATH];
	GetModuleFileName(_AtlBaseModule.GetResourceInstance(),path,_countof(path));
	*PathFindFileName(path)=0;
	wchar_t topic[_MAX_PATH];
	Sprintf(topic,_countof(topic),L"%s%sClassicShell.chm::/%s.html",path,GetDocRelativePath(),PathFindFileName(g_SettingsManager.GetRegPath()));
	HtmlHelp(GetDesktopWindow(),topic,HH_DISPLAY_TOPIC,NULL);
}

bool CSettingsDlg::IsTabValid( void )
{
	int idx=TabCtrl_GetCurSel(m_Tabs);
	if (idx<0) return true;
	TCITEM tab={TCIF_PARAM};
	TabCtrl_GetItem(m_Tabs,idx,&tab);
	CSetting *pGroup=(CSetting*)tab.lParam;
	ISettingsPanel *pPanel=pGroup->pPanel;
	if (!pPanel) pPanel=GetDefaultSettings();
	return pPanel->Validate(m_hWnd);
}

static CSettingsDlg g_SettingsDlg;

void EditSettings( const wchar_t *title, bool bModal )
{
	if (g_SettingsDlg.m_hWnd)
	{
		HWND top=GetWindow(g_SettingsDlg,GW_ENABLEDPOPUP);
		if (!top) top=g_SettingsDlg.m_hWnd;
		SetForegroundWindow(top);
		SetActiveWindow(top);
	}
	else
	{
		{
			CSettingsLockWrite lock;
			UpdateSettings();
		}
		DLGTEMPLATE *pTemplate=LoadDialogEx(IDD_SETTINGS);
		g_SettingsDlg.Init(g_SettingsManager.GetSettings());
		g_SettingsDlg.Create(NULL,pTemplate);
		g_SettingsDlg.SetWindowText(title);
		g_SettingsDlg.SetWindowPos(HWND_TOPMOST,0,0,0,0,SWP_NOSIZE|SWP_NOMOVE|(g_SettingsDlg.GetOnTop()?0:SWP_NOZORDER)|SWP_SHOWWINDOW);
		SetForegroundWindow(g_SettingsDlg.m_hWnd);
		if (bModal)
		{
			MSG msg;
			while (g_SettingsDlg.m_hWnd && GetMessage(&msg,0,0,0))
			{
				if (IsSettingsMessage(&msg)) continue;
				TranslateMessage(&msg);
				DispatchMessage(&msg);
			}
		}
	}
}

void CloseSettings( void )
{
	if (g_SettingsDlg.m_hWnd)
		g_SettingsDlg.DestroyWindow();
}

// Process the dialog messages for the settings box
bool IsSettingsMessage( MSG *msg )
{
	if (!g_SettingsDlg) return false;
	if (msg->hwnd!=g_SettingsDlg && !IsChild(g_SettingsDlg,msg->hwnd)) return false;
	// only process keyboard messages. if we process all messages the settings box gets stuck. I don't know why.
	if (msg->message<WM_KEYFIRST || msg->message>WM_KEYLAST) return false;

	// don't process any messages if a menu is up
	GUITHREADINFO threadInfo={sizeof(threadInfo)};
	GetGUIThreadInfo(GetCurrentThreadId(),&threadInfo);
	if (threadInfo.flags&GUI_INMENUMODE) return false;

	// handle Ctrl+Tab and Ctrl+Shift+Tab
	if (msg->message==WM_KEYDOWN && msg->wParam==VK_TAB && GetKeyState(VK_CONTROL)<0)
	{
		g_SettingsDlg.SendMessage(WM_KEYDOWN,VK_TAB,msg->lParam);
		return true;
	}
	if (msg->message==WM_KEYDOWN && msg->wParam==VK_F1 && GetKeyState(VK_CONTROL)>=0 && GetKeyState(VK_SHIFT)>=0 && GetKeyState(VK_MENU)>=0)
	{
		g_SettingsDlg.ShowHelp();
	}
	return IsDialogMessage(g_SettingsDlg,msg)!=0;
}

///////////////////////////////////////////////////////////////////////////////

void InitSettings( CSetting *pSettings, TSettingsComponent component )
{
	g_SettingsManager.Init(pSettings,component);
}

void LoadSettings( void )
{
	CSettingsLockWrite lock;
	g_SettingsManager.LoadSettings();
}

void UpdateDefaultSettings( void )
{
	if (!g_SettingsDlg.m_hWnd)
		UpdateSettings();
}

bool GetSettingBool( const wchar_t *name )
{
	return g_SettingsManager.GetSettingBool(name);
}

int GetSettingInt( const wchar_t *name )
{
	return g_SettingsManager.GetSettingInt(name);
}

int GetSettingInt( const wchar_t *name, bool &bDef )
{
	return g_SettingsManager.GetSettingInt(name,bDef);
}

bool GetSettingBool( const wchar_t *name, bool &bDef )
{
	return g_SettingsManager.GetSettingBool(name,bDef);
}

CString GetSettingString( const wchar_t *name )
{
	return g_SettingsManager.GetSettingString(name);
}

HIMAGELIST GetSettingsImageList( HWND tree )
{
	return g_SettingsManager.GetImageList(tree);
}

const wchar_t *GetSettingsRegPath( void )
{
	return g_SettingsManager.GetRegPath();
}

// Updates the setting with a new default value, and locked/hidden flags
void UpdateSetting( const wchar_t *name, const CComVariant &defValue, bool bLockedGP, bool bHidden )
{
	ATLASSERT(g_LockState==2); // must be locked for writing
	for (CSetting *pSetting=g_SettingsManager.GetSettings();pSetting->name;pSetting++)
		if (pSetting->type>=0 && wcscmp(pSetting->name,name)==0)
		{
			if (bLockedGP)
				pSetting->flags|=CSetting::FLAG_LOCKED_GP|CSetting::FLAG_DEFAULT;
			else
				pSetting->flags&=~CSetting::FLAG_LOCKED_GP;
			if (bHidden)
				pSetting->flags|=CSetting::FLAG_HIDDEN;
			else
				pSetting->flags&=~CSetting::FLAG_HIDDEN;
			pSetting->defValue=defValue;
			if (pSetting->flags&CSetting::FLAG_DEFAULT)
				pSetting->value=defValue;
			return;
		}
	ATLASSERT(0);
}

// Updates the setting with a new tooltip and a warning flag
void UpdateSetting( const wchar_t *name, int tipID, bool bWarning )
{
	ATLASSERT(g_LockState==2); // must be locked for writing
	for (CSetting *pSetting=g_SettingsManager.GetSettings();pSetting->name;pSetting++)
		if (pSetting->type>=0 && wcscmp(pSetting->name,name)==0)
		{
			if (bWarning)
				pSetting->flags|=CSetting::FLAG_WARNING;
			else
				pSetting->flags&=~CSetting::FLAG_WARNING;
			pSetting->tipID=tipID;
			return;
		}
	ATLASSERT(0);
}

void HideSettingGroup( const wchar_t *name )
{
	ATLASSERT(g_LockState==2); // must be locked for writing
	for (CSetting *pSetting=g_SettingsManager.GetSettings();pSetting->name;pSetting++)
		if (pSetting->type==CSetting::TYPE_GROUP && wcscmp(pSetting->name,name)==0)
		{
			pSetting->flags|=CSetting::FLAG_HIDDEN;
			return;
		}
	ATLASSERT(0);
}

const CSetting *GetAllSettings( void )
{
	ATLASSERT(g_bUIThread);
	return g_SettingsManager.GetSettings();
}
