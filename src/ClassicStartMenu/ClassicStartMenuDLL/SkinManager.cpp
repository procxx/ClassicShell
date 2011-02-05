// Classic Shell (c) 2009-2010, Ivo Beltchev
// The sources for Classic Shell are distributed under the MIT open source license

#include "stdafx.h"
#include "resource.h"
#include "SkinManager.h"
#include "IconManager.h"
#include "LogManager.h"
#include "SettingsParser.h"
#include "Settings.h"
#include "Translations.h"
#include "ResourceHelper.h"
#include "FNVHash.h"
#include "dllmain.h"
#include <dwmapi.h>

wchar_t MenuSkin::s_SkinError[1024];

const RECT DEFAULT_ICON_PADDING={3,3,3,3};
const RECT DEFAULT_TEXT_PADDING={1,0,4,0};
const int DEFAULT_ARROW_SIZE=4;
const SIZE DEFAULT_ARROW_PADDING={8,4};

MenuSkin::MenuSkin( void )
{
	AboutIcon=NULL;
	Main_bitmap=NULL;
	Caption_font=NULL;
	User_font=NULL;
	Main_font=NULL;
	Main_font2=NULL;
	Main_selectionColor=true;
	Main_selectionColor2=true;
	Main_separator=NULL;
	Main_separator2=NULL;
	Main_separatorV=NULL;
	Main_arrow=NULL;
	Main_arrow2=NULL;
	Main_pager=NULL;
	Main_pager_arrows=NULL;
	User_bitmap=NULL;
	Submenu_bitmap=NULL;
	Submenu_font=NULL;
	Submenu_selectionColor=true;
	Submenu_separator=NULL;
	Submenu_arrow=NULL;
	Submenu_separatorV=NULL;
	Submenu_pager=NULL;
	Submenu_pager_arrows=NULL;
	Search_bitmap=NULL;
}

MenuSkin::~MenuSkin( void )
{
	Reset();
}

void MenuSkin::Reset( void )
{
	if (AboutIcon) DestroyIcon(AboutIcon);
	if (Main_bitmap) DeleteObject(Main_bitmap);
	if (Caption_font) DeleteObject(Caption_font);
	if (User_font) DeleteObject(User_font);
	if (Main_font) DeleteObject(Main_font);
	if (Main_font2 && Main_font2!=Main_font) DeleteObject(Main_font2);
	if (!Main_selectionColor && Main_selection.bmp) DeleteObject(Main_selection.bmp);
	if (!Main_selectionColor2 && Main_selection2.bmp && (Main_selectionColor || Main_selection2.bmp!=Main_selection.bmp)) DeleteObject(Main_selection2.bmp);
	if (Main_separator) DeleteObject(Main_separator);
	if (Main_separator2 && Main_separator2!=Main_separator) DeleteObject(Main_separator2);
	if (Main_separatorV) DeleteObject(Main_separatorV);
	if (Main_arrow) DeleteObject(Main_arrow);
	if (Main_arrow2 && Main_arrow2!=Main_arrow) DeleteObject(Main_arrow2);
	if (User_bitmap) DeleteObject(User_bitmap);
	if (Submenu_bitmap) DeleteObject(Submenu_bitmap);
	if (Submenu_font) DeleteObject(Submenu_font);
	if (!Submenu_selectionColor && Submenu_selection.bmp) DeleteObject(Submenu_selection.bmp);
	if (Submenu_separator) DeleteObject(Submenu_separator);
	if (Submenu_separatorV) DeleteObject(Submenu_separatorV);
	if (Submenu_arrow) DeleteObject(Submenu_arrow);
	if (Main_pager) DeleteObject(Main_pager);
	if (Main_pager_arrows) DeleteObject(Main_pager_arrows);
	if (Submenu_pager) DeleteObject(Submenu_pager);
	if (Submenu_pager_arrows) DeleteObject(Submenu_pager_arrows);
	if (Search_bitmap) DeleteObject(Search_bitmap);

	AboutIcon=NULL;
	Main_bitmap=NULL;
	Caption_font=NULL;
	User_font=NULL;
	Main_font=NULL;
	Main_font2=NULL;
	Main_selectionColor=true;
	Main_selectionColor2=true;
	Main_separator=NULL;
	Main_separator2=NULL;
	Main_separatorV=NULL;
	Main_arrow=NULL;
	Main_arrow2=NULL;
	User_bitmap=NULL;
	Submenu_bitmap=NULL;
	Submenu_font=NULL;
	Submenu_selectionColor=true;
	Submenu_separator=NULL;
	Submenu_separatorV=NULL;
	Submenu_arrow=NULL;
	Main_pager=NULL;
	Main_pager_arrows=NULL;
	Submenu_pager=NULL;
	Submenu_pager_arrows=NULL;
	Search_bitmap=NULL;
}

static void GetErrorMessage( wchar_t *err, int size, DWORD code )
{
	FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM|FORMAT_MESSAGE_IGNORE_INSERTS,NULL,code,0,err,size,NULL);
}

static void LoadSkinNumbers( const wchar_t *str, int *numbers, int count, bool bColors )
{
	for (int i=0;i<count;i++)
	{
		wchar_t token[256];
		str=GetToken(str,token,_countof(token),L", \t");
		wchar_t *end;
		int num;
		if (token[0]=='#')
			num=wcstol(token+1,&end,16);
		else
			num=wcstol(token,&end,10);
		if (bColors)
			numbers[i]=((num&0xFF)<<16)|(num&0xFF00)|((num>>16)&0xFF);
		else
			numbers[i]=num;
	}
}

static HFONT LoadSkinFont( const wchar_t *str, const wchar_t *name, int weight, int size, bool bDPI )
{
	DWORD quality=DEFAULT_QUALITY;
	int smoothing=GetSettingInt(L"FontSmoothing");
	if (smoothing==1)
		quality=NONANTIALIASED_QUALITY;
	else if (smoothing==2)
		quality=ANTIALIASED_QUALITY;
	if (smoothing==3)
		quality=CLEARTYPE_QUALITY;
	wchar_t token[256];
	bool bItalic=false;
	if (str)
	{
		str=GetToken(str,token,_countof(token),L", \t");
		name=token;
		wchar_t token2[256];
		str=GetToken(str,token2,_countof(token2),L", \t");
		weight=FW_NORMAL;
		if (_wcsicmp(token2,L"bold")==0)
			weight=FW_BOLD;
		else if (_wcsicmp(token2,L"italic")==0)
			bItalic=true;
		else if (_wcsicmp(token2,L"bold_italic")==0)
			weight=FW_BOLD, bItalic=true;
		str=GetToken(str,token2,_countof(token2),L", \t");
		size=_wtol(token2);
	}
	else if (!name)
	{
		// get the default menu font
		NONCLIENTMETRICS metrics={sizeof(metrics)};
		SystemParametersInfo(SPI_GETNONCLIENTMETRICS,NULL,&metrics,0);
		metrics.lfMenuFont.lfQuality=(BYTE)quality;
		return CreateFontIndirect(&metrics.lfMenuFont);
	}
	int dpi=bDPI?CIconManager::GetDPI():96;
	return CreateFont(size*dpi/72,0,0,0,weight,bItalic?1:0,0,0,DEFAULT_CHARSET,OUT_DEFAULT_PRECIS,CLIP_DEFAULT_PRECIS,quality,DEFAULT_PITCH,name);
}

static HICON LoadSkinIcon( HMODULE hMod, int index )
{
	if (hMod)
	{
		return (HICON)LoadImage(hMod,MAKEINTRESOURCE(index),IMAGE_ICON,0,0,LR_DEFAULTSIZE);
	}
	else
	{
		wchar_t path[_MAX_PATH];
		GetSkinsPath(path);
		wchar_t fname[_MAX_PATH];
		Sprintf(fname,_countof(fname),L"%s%d.ico",path,index);
		return (HICON)LoadImage(NULL,fname,IMAGE_ICON,0,0,LR_DEFAULTSIZE|LR_LOADFROMFILE);
	}
}

static HBITMAP LoadSkinBitmap( HMODULE hMod, int index, int maskIndex, bool &b32, COLORREF menuColor )
{
	wchar_t err[1024];
	HBITMAP bmp;
	if (hMod)
	{
		bmp=(HBITMAP)LoadImage(hMod,MAKEINTRESOURCE(index),IMAGE_BITMAP,0,0,LR_CREATEDIBSECTION);
		if (!bmp)
		{
			GetErrorMessage(err,_countof(err),GetLastError());
			Sprintf(MenuSkin::s_SkinError,_countof(MenuSkin::s_SkinError),LoadStringEx(IDS_SKIN_ERR_BMPRES),index,err);
			return NULL;
		}
	}
	else
	{
		wchar_t path[_MAX_PATH];
		GetSkinsPath(path);
		wchar_t fname[_MAX_PATH];
		Sprintf(fname,_countof(fname),L"%s%d.bmp",path,index);
		bmp=(HBITMAP)LoadImage(NULL,fname,IMAGE_BITMAP,0,0,LR_CREATEDIBSECTION|LR_LOADFROMFILE);
		if (!bmp)
		{
			GetErrorMessage(err,_countof(err),GetLastError());
			Sprintf(MenuSkin::s_SkinError,_countof(MenuSkin::s_SkinError),LoadStringEx(IDS_SKIN_ERR_BMPFILE),fname,err);
			return NULL;
		}
	}


	HBITMAP bmpMask=NULL;
	BITMAP infoMask;
	if (maskIndex>0)
	{
		if (hMod)
		{
			bmpMask=(HBITMAP)LoadImage(hMod,MAKEINTRESOURCE(maskIndex),IMAGE_BITMAP,0,0,LR_CREATEDIBSECTION);
			if (!bmpMask)
			{
				GetErrorMessage(err,_countof(err),GetLastError());
				Sprintf(MenuSkin::s_SkinError,_countof(MenuSkin::s_SkinError),LoadStringEx(IDS_SKIN_ERR_MASKRES),maskIndex,err);
			}
		}
		else
		{
			wchar_t path[_MAX_PATH];
			GetSkinsPath(path);
			wchar_t fname[_MAX_PATH];
			Sprintf(fname,_countof(fname),L"%s%d.bmp",path,maskIndex);
			bmpMask=(HBITMAP)LoadImage(NULL,fname,IMAGE_BITMAP,0,0,LR_CREATEDIBSECTION|LR_LOADFROMFILE);
			if (!bmpMask)
			{
				GetErrorMessage(err,_countof(err),GetLastError());
				Sprintf(MenuSkin::s_SkinError,_countof(MenuSkin::s_SkinError),LoadStringEx(IDS_SKIN_ERR_MASKFILE),fname,err);
			}
		}
		if (bmpMask)
			GetObject(bmpMask,sizeof(infoMask),&infoMask);
	}

	BITMAP info;
	GetObject(bmp,sizeof(info),&info);

	if (bmpMask && (info.bmWidth!=infoMask.bmWidth || info.bmHeight!=infoMask.bmHeight))
	{
		Sprintf(MenuSkin::s_SkinError,_countof(MenuSkin::s_SkinError),LoadStringEx(IDS_SKIN_ERR_MASKSIZE),index,maskIndex);
	}

	if (bmpMask && info.bmWidth==infoMask.bmWidth && info.bmHeight==infoMask.bmHeight)
	{
		// apply color mask
		unsigned char *ptr=(unsigned char*)info.bmBits;
		int stride=info.bmBitsPixel/8;
		int pad=stride*info.bmWidth;
		pad=((pad+3)&~3)-pad;
		unsigned char *ptrMask=(unsigned char*)infoMask.bmBits;
		int strideMask=infoMask.bmBitsPixel/8;
		int padMask=strideMask*info.bmWidth;
		padMask=((padMask+3)&~3)-padMask;

		int dr=0, dg=0, db=0, da=0;
		int mr=(menuColor)&255, mg=(menuColor>>8)&255, mb=(menuColor>>16)&255;

		DWORD dwmColor;
		BOOL dwmOpaque;
		if (SUCCEEDED(DwmGetColorizationColor(&dwmColor,&dwmOpaque)))
		{
			da=(dwmColor>>24)&255;
			dr=(dwmColor>>16)&255;
			dg=(dwmColor>>8)&255;
			db=(dwmColor)&255;
		}

		for (int y=0;y<info.bmHeight;y++)
		{
			for (int x=0;x<info.bmWidth;x++,ptr+=stride,ptrMask+=strideMask)
			{
				int a1=(ptrMask[2]*da)/255, a2=ptrMask[1];
				int a3=255-a1-a2; if (a3<0) a3=0;
				int b=(db*a1+mb*a2+ptr[0]*a3)/255; if (b>255) b=255;
				int g=(dg*a1+mg*a2+ptr[1]*a3)/255; if (g>255) g=255;
				int r=(dr*a1+mr*a2+ptr[2]*a3)/255; if (r>255) r=255;
				ptr[0]=(unsigned char)b;
				ptr[1]=(unsigned char)g;
				ptr[2]=(unsigned char)r;
			}
			ptr+=pad;
			ptrMask+=padMask;
		}
	}
	if (bmpMask) DeleteObject(bmpMask);

	int n=info.bmWidth*info.bmHeight;
	b32=false;
	if (info.bmBitsPixel<32)
		return bmp;

	// HACK: when LoadImage reads a 24-bit image it creates a 32-bit bitmap with 0 in the alpha channel
	// we use that to detect 24-bit images and don't pre-multiply the alpha
	for (int i=0;i<n;i++)
	{
		unsigned int &pixel=((unsigned int*)info.bmBits)[i];
		if (pixel&0xFF000000)
		{
			b32=true;
			break;
		}
	}

	if (!b32) return bmp;

	// 32-bit bitmap detected. pre-multiply the alpha
	for (int i=0;i<n;i++)
	{
		unsigned int &pixel=((unsigned int*)info.bmBits)[i];
		int a=(pixel>>24);
		int r=(pixel>>16)&255;
		int g=(pixel>>8)&255;
		int b=(pixel)&255;
		r=(r*a)/255;
		g=(g*a)/255;
		b=(b*a)/255;
		pixel=(a<<24)|(r<<16)|(g<<8)|b;
	}
	return bmp;
}

static void MirrorBitmap( HBITMAP bmp )
{
	BITMAP info;
	GetObject(bmp,sizeof(info),&info);

	unsigned char *ptr=(unsigned char*)info.bmBits;
	if (!ptr) return;
	int stride=info.bmBitsPixel/8;
	int pitch=stride*info.bmWidth;
	for (int y=0;y<info.bmHeight;y++,ptr+=pitch)
	{
		unsigned char *start=ptr;
		unsigned char *end=ptr+stride*(info.bmWidth-1);
		for (int x=0;x<info.bmWidth;x+=2,start+=stride,end-=stride)
		{
			char q[4];
			memcpy(q,start,stride);
			memcpy(start,end,stride);
			memcpy(end,q,stride);
		}
	}
}

static bool FindToken( const wchar_t *str, const wchar_t *token )
{
	wchar_t text[256];
	while (*str)
	{
		str=GetToken(str,text,_countof(text),L" \t,");
		if (_wcsicmp(text,token)==0)
			return true;
	}
	return false;
}

// Load the skin from the module. If hMod is NULL loads the "custom" skin from 1.txt
static bool LoadSkin( HMODULE hMod, MenuSkin &skin, const wchar_t *variation, const std::vector<unsigned int> &options, unsigned int flags )
{
	skin.version=1;
	CSkinParser parser;
	if (hMod)
	{
		HRSRC hResInfo=FindResource(hMod,MAKEINTRESOURCE(1),L"SKIN");
		if (!hResInfo)
		{
			Strcpy(MenuSkin::s_SkinError,_countof(MenuSkin::s_SkinError),LoadStringEx(IDS_SKIN_ERR_FIND_RES1));
			return false;
		}
		if (!parser.LoadText(hMod,hResInfo))
		{
			Strcpy(MenuSkin::s_SkinError,_countof(MenuSkin::s_SkinError),LoadStringEx(IDS_SKIN_ERR_LOAD_RES1));
			return false;
		}
	}
	else
	{
		wchar_t path[_MAX_PATH];
		GetSkinsPath(path);
		Strcat(path,_countof(path),L"1.txt");
		if (!parser.LoadText(path))
		{
			Sprintf(MenuSkin::s_SkinError,_countof(MenuSkin::s_SkinError),LoadStringEx(IDS_SKIN_ERR_LOAD_FILE1),path);
			return false;
		}
	}
	parser.ParseText();

	const wchar_t *str;

	CString AllPrograms_options;
	str=parser.FindSetting(L"AllPrograms_options");
	if (str)
		AllPrograms_options=str;
	if ((flags&LOADMENU_MAIN) || AllPrograms_options.IsEmpty() || FindToken(AllPrograms_options,L"variations"))
	{
		for (int i=1;;i++)
		{
			wchar_t name[20];
			Sprintf(name,_countof(name),L"Variation%d",i);
			str=parser.FindSetting(name);
			if (str)
			{
				LOG_MENU(LOG_OPEN,L"Variation setting: '%s'",str);
				wchar_t token[256];
				str=GetToken(str,token,_countof(token),L", \t");
				int  res=_wtol(token);
				str=GetToken(str,token,_countof(token),L", \t");
				skin.Variations.push_back(std::pair<int,CString>(res,token));
				LOG_MENU(LOG_OPEN,L"Variation found: name=%s, id=%d",token,res);
			}
			else
				break;
		}
	}

	if (variation)
	{
		for (std::vector<std::pair<int,CString>>::const_iterator it=skin.Variations.begin();it!=skin.Variations.end();++it)
			if (wcscmp(variation,it->second)==0)
			{
				if (it->first<=1) break;
				LOG_MENU(LOG_OPEN,L"Loading variation: name=%s, id=%d",it->second,it->first);
				if (hMod)
				{
					HRSRC hResInfo=FindResource(hMod,MAKEINTRESOURCE(it->first),L"SKIN");
					if (!hResInfo)
					{
						Sprintf(MenuSkin::s_SkinError,_countof(MenuSkin::s_SkinError),LoadStringEx(IDS_SKIN_ERR_FIND_RES),it->first);
						break;
					}
					if (!parser.LoadVariation(hMod,hResInfo))
					{
						Sprintf(MenuSkin::s_SkinError,_countof(MenuSkin::s_SkinError),LoadStringEx(IDS_SKIN_ERR_LOAD_RES),it->first);
						break;
					}
				}
				else
				{
					wchar_t path[_MAX_PATH];
					GetSkinsPath(path);
					wchar_t name[20];
					Sprintf(name,_countof(name),L"%d.txt",it->first);
					Strcat(path,_countof(path),name);
					if (!parser.LoadVariation(path))
					{
						Sprintf(MenuSkin::s_SkinError,_countof(MenuSkin::s_SkinError),LoadStringEx(IDS_SKIN_ERR_LOAD_FILE),path);
						break;
					}
				}

				break;
			}
	}

	// find options
	std::vector<const wchar_t*> values; // list of true values
	for (int i=0;;i++)
	{
		MenuSkin::Option option;
		if (!parser.ParseOption(option.name,option.label,option.value,option.condition,option.value2,i))
			break;
		if (!option.name.IsEmpty())
		{
			if (!(flags&LOADMENU_MAIN) && !AllPrograms_options.IsEmpty() && !FindToken(AllPrograms_options,option.name))
				continue;
			skin.Options.push_back(option);
			unsigned int hash=CalcFNVHash(option.name)&0xFFFFFFFE;
			bool bValue=option.value;
			for (std::vector<unsigned int>::const_iterator it=options.begin();it!=options.end();++it)
				if ((*it&0xFFFFFFFE)==hash)
				{
					bValue=(*it&1)!=0;
					break;
				}
			if (!option.condition.IsEmpty() && !EvalCondition(option.condition,values.empty()?NULL:&values[0],(int)values.size()))
				bValue=option.value2;
			if (bValue)
				values.push_back(option.name);
		}
	}
	if (!(flags&LOADMENU_MAIN))
		values.push_back(L"ALL_PROGRAMS");
	if (values.empty())
		parser.FilterConditions(NULL,0);
	else
		parser.FilterConditions(&values[0],(int)values.size());

	// parse settings
	str=parser.FindSetting(L"About");
	if (str)
	{
		skin.About=str;
		skin.About.Replace(L"\\n",L"\n");
	}
	else skin.About.Empty();

	str=parser.FindSetting(L"AboutIcon");
	if (str) skin.AboutIcon=LoadSkinIcon(hMod,_wtol(str));

	str=parser.FindSetting(L"Version");
	if (str)
		skin.version=_wtol(str);

	skin.ForceRTL=false;
	if (!hMod)
	{
		str=parser.FindSetting(L"ForceRTL");
		if (str && _wtol(str))
			skin.ForceRTL=true;
	}

	bool bRTL=skin.ForceRTL || IsLanguageRTL();

	// CAPTION SECTION - describes the caption portion of the main menu
	str=parser.FindSetting(L"Caption_font");
	skin.Caption_font=LoadSkinFont(str,L"Segoe UI",FW_NORMAL,18,true);
	str=parser.FindSetting(L"Caption_text_color");
	if (str)
		LoadSkinNumbers(str,(int*)&skin.Caption_text_color,1,true);
	else
		skin.Caption_text_color=0xFFFFFF;

	str=parser.FindSetting(L"Caption_glow_color");
	if (str)
		LoadSkinNumbers(str,(int*)&skin.Caption_glow_color,1,true);
	else
		skin.Caption_glow_color=0xFFFFFF;

	str=parser.FindSetting(L"Caption_glow_size");
	if (str)
		skin.Caption_glow_size=_wtol(str);
	else
		skin.Caption_glow_size=0;

	str=parser.FindSetting(L"Caption_padding");
	if (str)
		LoadSkinNumbers(str,(int*)&skin.Caption_padding,4,false);
	else
		memset(&skin.Caption_padding,0,sizeof(skin.Caption_padding));


	// MENU SECTION - describes the menu portion of the main menu
	str=parser.FindSetting(L"Main_background");
	if (str)
		LoadSkinNumbers(str,(int*)&skin.Main_background,1,true);
	else
		skin.Main_background=GetSysColor(COLOR_MENU);

	str=parser.FindSetting(L"Main_bitmap");
	if (str && flags==(LOADMENU_MAIN|LOADMENU_RESOURCES))
	{
		int id=_wtol(str);
		if (id)
		{
			str=parser.FindSetting(L"Main_bitmap_mask");
			int id2=str?_wtol(str):0;
			skin.Main_bitmap=LoadSkinBitmap(hMod,id,id2,skin.Main_bitmap32,skin.Main_background);
			if (!skin.Main_bitmap) return false;
		}
	}
	skin.bTwoColumns=false;
	str=parser.FindSetting(L"Main_2columns");
	if (str && _wtol(str))
		skin.bTwoColumns=true;
	str=parser.FindSetting(L"Main_bitmap_slices_X");
	memset(skin.Main_bitmap_slices_X,0,sizeof(skin.Main_bitmap_slices_X));
	if (str)
		LoadSkinNumbers(str,skin.Main_bitmap_slices_X+(skin.bTwoColumns?3:0),6,false);
	str=parser.FindSetting(L"Main_bitmap_slices_Y");
	if (str)
		LoadSkinNumbers(str,skin.Main_bitmap_slices_Y,_countof(skin.Main_bitmap_slices_Y),false);
	else
		memset(skin.Main_bitmap_slices_Y,0,sizeof(skin.Main_bitmap_slices_Y));
	str=parser.FindSetting(L"Main_opacity");
	skin.Main_opacity=MenuSkin::OPACITY_SOLID;
	if (str && skin.Main_bitmap)
	{
		if (_wcsicmp(str,L"region")==0) skin.Main_opacity=MenuSkin::OPACITY_REGION;
		if (_wcsicmp(str,L"alpha")==0) skin.Main_opacity=MenuSkin::OPACITY_ALPHA;
		if (_wcsicmp(str,L"glass")==0) skin.Main_opacity=MenuSkin::OPACITY_GLASS;
		if (_wcsicmp(str,L"fullalpha")==0) skin.Main_opacity=MenuSkin::OPACITY_FULLALPHA;
		if (_wcsicmp(str,L"fullglass")==0) skin.Main_opacity=MenuSkin::OPACITY_FULLGLASS;
	}
	str=parser.FindSetting(L"Main_opacity2");
	skin.Main_opacity2=skin.Main_opacity;
	if (str)
	{
		if (skin.Main_opacity2==MenuSkin::OPACITY_ALPHA || skin.Main_opacity2==MenuSkin::OPACITY_FULLALPHA)
		{
			if (_wcsicmp(str,L"alpha")==0) skin.Main_opacity2=MenuSkin::OPACITY_ALPHA;
			if (_wcsicmp(str,L"fullalpha")==0) skin.Main_opacity2=MenuSkin::OPACITY_FULLALPHA;
		}
		if (skin.Main_opacity2==MenuSkin::OPACITY_GLASS || skin.Main_opacity2==MenuSkin::OPACITY_FULLGLASS)
		{
			if (_wcsicmp(str,L"glass")==0) skin.Main_opacity2=MenuSkin::OPACITY_GLASS;
			if (_wcsicmp(str,L"fullglass")==0) skin.Main_opacity2=MenuSkin::OPACITY_FULLGLASS;
		}
	}
	str=parser.FindSetting(L"Main_large_icons");
	skin.Main_large_icons=str && _wtol(str);

	str=parser.FindSetting(L"Main_font");
	skin.Main_font=LoadSkinFont(str,NULL,0,0,true);

	str=parser.FindSetting(L"Main_font2");
	if (str)
		skin.Main_font2=LoadSkinFont(str,NULL,0,0,true);
	else
		skin.Main_font2=skin.Main_font;

	str=parser.FindSetting(L"Main_glow_size");
	if (str)
		skin.Main_glow_size=_wtol(str);
	else
		skin.Main_glow_size=0;

	str=parser.FindSetting(L"Main_glow_size2");
	if (str)
		skin.Main_glow_size2=_wtol(str);
	else
		skin.Main_glow_size2=skin.Main_glow_size;

	str=parser.FindSetting(L"Main_text_color");
	if (str)
		LoadSkinNumbers(str,(int*)skin.Main_text_color,_countof(skin.Main_text_color),true);
	else
	{
		skin.Main_text_color[0]=GetSysColor(COLOR_MENUTEXT);
		skin.Main_text_color[1]=GetSysColor(COLOR_HIGHLIGHTTEXT);
		skin.Main_text_color[2]=GetSysColor(COLOR_GRAYTEXT);
		skin.Main_text_color[3]=GetSysColor(COLOR_HIGHLIGHTTEXT);
	}
	str=parser.FindSetting(L"Main_text_color2");
	if (str)
		LoadSkinNumbers(str,(int*)skin.Main_text_color2,_countof(skin.Main_text_color2),true);
	else
		memcpy(skin.Main_text_color2,skin.Main_text_color,sizeof(skin.Main_text_color2));
	str=parser.FindSetting(L"Main_arrow_color");
	if (str)
		LoadSkinNumbers(str,(int*)skin.Main_arrow_color,_countof(skin.Main_arrow_color),true);
	else
	{
		skin.Main_arrow_color[0]=skin.Main_text_color[0];
		skin.Main_arrow_color[1]=skin.Main_text_color[1];
	}
	str=parser.FindSetting(L"Main_arrow_color2");
	if (str)
		LoadSkinNumbers(str,(int*)skin.Main_arrow_color2,_countof(skin.Main_arrow_color2),true);
	else
		memcpy(skin.Main_arrow_color2,skin.Main_arrow_color,sizeof(skin.Main_arrow_color2));

	str=parser.FindSetting(L"Main_padding");
	if (str)
		LoadSkinNumbers(str,(int*)&skin.Main_padding,4,false);
	else
		memset(&skin.Main_padding,0,sizeof(skin.Main_padding));

	str=parser.FindSetting(L"Main_padding2");
	if (str)
		LoadSkinNumbers(str,(int*)&skin.Main_padding2,4,false);
	else
		memset(&skin.Main_padding2,-1,sizeof(skin.Main_padding2));

	str=parser.FindSetting(L"Main_selection");
	if (str)
	{
		if (str[0]=='#')
			LoadSkinNumbers(str,(int*)&skin.Main_selection.color,1,true);
		else
		{
			skin.Main_selectionColor=false;
			skin.Main_selection.bmp=NULL;
			if (flags==(LOADMENU_MAIN|LOADMENU_RESOURCES))
			{
				int id=_wtol(str);
				if (id)
				{
					skin.Main_selection.bmp=LoadSkinBitmap(hMod,id,0,skin.Main_selection32,0);
					if (!skin.Main_selection.bmp) return false;
				}
			}

			str=parser.FindSetting(L"Main_selection_slices_X");
			if (str)
				LoadSkinNumbers(str,(int*)&skin.Main_selection_slices_X,_countof(skin.Main_selection_slices_X),false);
			else
				memset(skin.Main_selection_slices_X,0,sizeof(skin.Main_selection_slices_X));
		}
		str=parser.FindSetting(L"Main_selection_slices_Y");
		if (str)
			LoadSkinNumbers(str,(int*)&skin.Main_selection_slices_Y,_countof(skin.Main_selection_slices_Y),false);
		else
			memset(skin.Main_selection_slices_Y,0,sizeof(skin.Main_selection_slices_Y));
	}
	else
	{
		skin.Main_selection.color=GetSysColor(COLOR_MENUHILIGHT);
	}

	str=parser.FindSetting(L"Main_selection2");
	if (str)
	{
		if (str[0]=='#')
			LoadSkinNumbers(str,(int*)&skin.Main_selection2.color,1,true);
		else
		{
			skin.Main_selectionColor2=false;
			skin.Main_selection2.bmp=NULL;
			if (flags==(LOADMENU_MAIN|LOADMENU_RESOURCES))
			{
				int id=_wtol(str);
				if (id)
				{
					skin.Main_selection2.bmp=LoadSkinBitmap(hMod,id,0,skin.Main_selection232,0);
					if (!skin.Main_selection2.bmp) return false;
				}
			}

			str=parser.FindSetting(L"Main_selection_slices_X2");
			if (str)
				LoadSkinNumbers(str,(int*)&skin.Main_selection_slices_X2,_countof(skin.Main_selection_slices_X2),false);
			else
				memset(skin.Main_selection_slices_X2,0,sizeof(skin.Main_selection_slices_X2));
		}
		str=parser.FindSetting(L"Main_selection_slices_Y2");
		if (str)
			LoadSkinNumbers(str,(int*)&skin.Main_selection_slices_Y2,_countof(skin.Main_selection_slices_Y2),false);
		else
			memset(skin.Main_selection_slices_Y2,0,sizeof(skin.Main_selection_slices_Y2));
	}
	else
	{
		skin.Main_selectionColor2=skin.Main_selectionColor;
		if (skin.Main_selectionColor2)
			skin.Main_selection2.color=skin.Main_selection.color;
		else
		{
			skin.Main_selection2.bmp=skin.Main_selection.bmp;
			skin.Main_selection232=skin.Main_selection32;
			memcpy(skin.Main_selection_slices_X2,skin.Main_selection_slices_X,sizeof(skin.Main_selection_slices_X2));
			memcpy(skin.Main_selection_slices_Y2,skin.Main_selection_slices_Y,sizeof(skin.Main_selection_slices_Y2));
		}
	}

	if (bRTL && !skin.Main_selectionColor && skin.Main_selection.bmp)
	{
		MirrorBitmap(skin.Main_selection.bmp);
		int q=skin.Main_selection_slices_X[0]; skin.Main_selection_slices_X[0]=skin.Main_selection_slices_X[2]; skin.Main_selection_slices_X[2]=q;
	}

	if (bRTL && !skin.Main_selectionColor2 && (skin.Main_selectionColor || skin.Main_selection2.bmp!=skin.Main_selection.bmp))
	{
		MirrorBitmap(skin.Main_selection2.bmp);
		int q=skin.Main_selection_slices_X2[0]; skin.Main_selection_slices_X2[0]=skin.Main_selection_slices_X2[2]; skin.Main_selection_slices_X2[2]=q;
	}

	skin.Main_arrow_Size.cx=skin.Main_arrow_Size2.cx=DEFAULT_ARROW_SIZE;
	skin.Main_arrow_Size.cy=skin.Main_arrow_Size2.cy=0;
	if (flags==(LOADMENU_MAIN|LOADMENU_RESOURCES))
	{
		str=parser.FindSetting(L"Main_arrow");
		if (str)
		{
			int id=_wtol(str);
			if (id)
			{
				skin.Main_arrow=LoadSkinBitmap(hMod,id,0,skin.Main_arrow32,0);
				if (!skin.Main_arrow) return false;
				if (bRTL)
					MirrorBitmap(skin.Main_arrow);
				BITMAP info;
				GetObject(skin.Main_arrow,sizeof(info),&info);
				skin.Main_arrow_Size.cx=info.bmWidth;
				skin.Main_arrow_Size.cy=info.bmHeight/2;
			}
		}

		str=parser.FindSetting(L"Main_arrow2");
		if (str)
		{
			int id=_wtol(str);
			if (id)
			{
				skin.Main_arrow2=LoadSkinBitmap(hMod,id,0,skin.Main_arrow232,0);
				if (!skin.Main_arrow2) return false;
				if (bRTL)
					MirrorBitmap(skin.Main_arrow2);
				BITMAP info;
				GetObject(skin.Main_arrow2,sizeof(info),&info);
				skin.Main_arrow_Size2.cx=info.bmWidth;
				skin.Main_arrow_Size2.cy=info.bmHeight/2;
			}
		}
		else
		{
			skin.Main_arrow2=skin.Main_arrow;
			skin.Main_arrow232=skin.Main_arrow32;
			skin.Main_arrow_Size2=skin.Main_arrow_Size;
		}
	}
	str=parser.FindSetting(L"Main_arrow_padding");
	if (str)
		LoadSkinNumbers(str,(int*)&skin.Main_arrow_padding,2,false);
	else
		skin.Main_arrow_padding=DEFAULT_ARROW_PADDING;

	str=parser.FindSetting(L"Main_arrow_padding2");
	if (str)
		LoadSkinNumbers(str,(int*)&skin.Main_arrow_padding2,2,false);
	else
		skin.Main_arrow_padding2=skin.Main_arrow_padding;

	str=parser.FindSetting(L"Main_thin_frame");
	skin.Main_thin_frame=(str && _wtol(str));

	skin.Main_separatorHeight=skin.Main_separatorHeight2=0;
	if (flags==(LOADMENU_MAIN|LOADMENU_RESOURCES))
	{
		str=parser.FindSetting(L"Main_separator");
		if (str)
		{
			int id=_wtol(str);
			if (id)
			{
				skin.Main_separator=LoadSkinBitmap(hMod,id,0,skin.Main_separator32,0);
				if (!skin.Main_separator) return false;
				BITMAP info;
				GetObject(skin.Main_separator,sizeof(info),&info);
				skin.Main_separatorHeight=info.bmHeight;
				str=parser.FindSetting(L"Main_separator_slices_X");
				if (str)
				{
					LoadSkinNumbers(str,skin.Main_separator_slices_X,_countof(skin.Main_separator_slices_X),false);
				}
				else
					memset(skin.Main_separator_slices_X,0,sizeof(skin.Main_separator_slices_X));
			}
		}

		str=parser.FindSetting(L"Main_separator2");
		if (str)
		{
			int id=_wtol(str);
			if (id)
			{
				skin.Main_separator2=LoadSkinBitmap(hMod,id,0,skin.Main_separator232,0);
				if (!skin.Main_separator2) return false;
				BITMAP info;
				GetObject(skin.Main_separator2,sizeof(info),&info);
				skin.Main_separatorHeight2=info.bmHeight;
				str=parser.FindSetting(L"Main_separator_slices_X2");
				if (str)
				{
					LoadSkinNumbers(str,skin.Main_separator_slices_X2,_countof(skin.Main_separator_slices_X2),false);
				}
				else
					memset(skin.Main_separator_slices_X2,0,sizeof(skin.Main_separator_slices_X2));
			}
		}
		else
		{
			skin.Main_separator2=skin.Main_separator;
			skin.Main_separator232=skin.Main_separator32;
			skin.Main_separatorHeight2=skin.Main_separatorHeight;
			memcpy(skin.Main_separator_slices_X2,skin.Main_separator_slices_X,sizeof(skin.Main_separator_slices_X2));
		}

		str=parser.FindSetting(L"Main_separatorV");
		if (str)
		{
			int id=_wtol(str);
			if (id)
			{
				skin.Main_separatorV=LoadSkinBitmap(hMod,id,0,skin.Main_separatorV32,0);
				if (!skin.Main_separatorV) return false;
				BITMAP info;
				GetObject(skin.Main_separatorV,sizeof(info),&info);
				skin.Main_separatorWidth=info.bmWidth;
				str=parser.FindSetting(L"Main_separator_slices_Y");
				if (str)
				{
					LoadSkinNumbers(str,skin.Main_separator_slices_Y,_countof(skin.Main_separator_slices_Y),false);
				}
				else
					memset(skin.Main_separator_slices_Y,0,sizeof(skin.Main_separator_slices_Y));
			}
		}
	}

	str=parser.FindSetting(L"Main_icon_padding");
	if (str)
		LoadSkinNumbers(str,(int*)&skin.Main_icon_padding,4,false);
	else
		skin.Main_icon_padding=DEFAULT_ICON_PADDING;
	str=parser.FindSetting(L"Main_text_padding");
	if (str)
		LoadSkinNumbers(str,(int*)&skin.Main_text_padding,4,false);
	else
		skin.Main_text_padding=DEFAULT_TEXT_PADDING;

	str=parser.FindSetting(L"Main_icon_padding2");
	if (str)
		LoadSkinNumbers(str,(int*)&skin.Main_icon_padding2,4,false);
	else
		skin.Main_icon_padding2=skin.Main_icon_padding;
	str=parser.FindSetting(L"Main_text_padding2");
	if (str)
		LoadSkinNumbers(str,(int*)&skin.Main_text_padding2,4,false);
	else
		skin.Main_text_padding2=skin.Main_text_padding;

	str=parser.FindSetting(L"Main_pager");
	if (str && flags==(LOADMENU_MAIN|LOADMENU_RESOURCES))
	{
		int id=_wtol(str);
		if (id)
			skin.Main_pager=LoadSkinBitmap(hMod,id,0,skin.Main_pager32,0);
	}
	str=parser.FindSetting(L"Main_pager_slices_X");
	if (str)
		LoadSkinNumbers(str,skin.Main_pager_slices_X,_countof(skin.Main_pager_slices_X),false);
	else
		memset(skin.Main_pager_slices_X,0,sizeof(skin.Main_pager_slices_X));
	str=parser.FindSetting(L"Main_pager_slices_Y");
	if (str)
		LoadSkinNumbers(str,skin.Main_pager_slices_Y,_countof(skin.Main_pager_slices_Y),false);
	else
		memset(skin.Main_pager_slices_Y,0,sizeof(skin.Main_pager_slices_Y));
	if (bRTL && skin.Main_pager)
	{
		MirrorBitmap(skin.Main_pager);
		int q=skin.Main_pager_slices_X[0]; skin.Main_pager_slices_X[0]=skin.Main_pager_slices_X[2]; skin.Main_pager_slices_X[2]=q;
	}

	str=parser.FindSetting(L"Main_pager_arrows");
	if (str && flags==(LOADMENU_MAIN|LOADMENU_RESOURCES))
	{
		int id=_wtol(str);
		if (id)
		{
			skin.Main_pager_arrows=LoadSkinBitmap(hMod,id,0,skin.Main_pager_arrows32,0);
			BITMAP info;
			GetObject(skin.Main_pager_arrows,sizeof(info),&info);
			skin.Main_pager_arrow_Size.cx=info.bmWidth/2;
			skin.Main_pager_arrow_Size.cy=info.bmHeight/2;
		}
	}
	if (bRTL && skin.Main_pager_arrows)
		MirrorBitmap(skin.Main_pager_arrows);

	str=parser.FindSetting(L"User_bitmap");
	if (str && flags==(LOADMENU_MAIN|LOADMENU_RESOURCES))
	{
		int id=_wtol(str);
		if (id)
		{
			bool b32;
			skin.User_bitmap=LoadSkinBitmap(hMod,id,0,b32,0);
			if (!skin.User_bitmap) return false;
			if (bRTL)
				MirrorBitmap(skin.User_bitmap);
		}
	}

	str=parser.FindSetting(L"User_image_size");
	if (str)
	{
		skin.User_image_size=_wtol(str);
		if (skin.User_image_size<0) skin.User_image_size=0;
		if (skin.User_image_size>128) skin.User_image_size=128;
	}
	else
		skin.User_image_size=0;

	str=parser.FindSetting(L"User_frame_position");
	if (str)
	{
		wchar_t token[256];
		str=GetToken(str,token,_countof(token),L", \t");
		if (_wcsicmp(token,L"center")==0)
			skin.User_frame_position.x=MenuSkin::USER_CENTER;
		else if (_wcsicmp(token,L"center1")==0)
			skin.User_frame_position.x=MenuSkin::USER_CENTER1;
		else if (_wcsicmp(token,L"center2")==0)
			skin.User_frame_position.x=MenuSkin::USER_CENTER2;
		else
			skin.User_frame_position.x=_wtol(token);

		GetToken(str,token,_countof(token),L", \t");
		skin.User_frame_position.y=_wtol(token);
	}
	else
		memset(&skin.User_frame_position,0,sizeof(skin.User_frame_position));

	str=parser.FindSetting(L"User_image_offset");
	if (str && skin.User_bitmap)
		LoadSkinNumbers(str,(int*)&skin.User_image_offset,2,false);
	else
		skin.User_image_offset.x=skin.User_image_offset.y=2;

	str=parser.FindSetting(L"User_image_alpha");
	if (str)
		skin.User_image_alpha=_wtol(str)&255;
	else
		skin.User_image_alpha=255;

	str=parser.FindSetting(L"User_name_position");
	if (str)
		LoadSkinNumbers(str,(int*)&skin.User_name_position,4,false);
	else
		memset(&skin.User_name_position,0,sizeof(skin.User_name_position));

	skin.User_name_align=MenuSkin::NAME_CENTER;
	str=parser.FindSetting(L"User_name_align");
	if (str)
	{
		if (_wcsicmp(str,L"center1")==0)
			skin.User_name_align=MenuSkin::NAME_CENTER1;
		else if (_wcsicmp(str,L"center2")==0)
			skin.User_name_align=MenuSkin::NAME_CENTER2;
		else if (_wcsicmp(str,L"left")==0)
			skin.User_name_align=MenuSkin::NAME_LEFT;
		else if (_wcsicmp(str,L"left1")==0)
			skin.User_name_align=MenuSkin::NAME_LEFT1;
		else if (_wcsicmp(str,L"left2")==0)
			skin.User_name_align=MenuSkin::NAME_LEFT2;
		else if (_wcsicmp(str,L"right")==0)
			skin.User_name_align=MenuSkin::NAME_RIGHT;
		else if (_wcsicmp(str,L"right1")==0)
			skin.User_name_align=MenuSkin::NAME_RIGHT1;
		else if (_wcsicmp(str,L"right2")==0)
			skin.User_name_align=MenuSkin::NAME_RIGHT2;
	}

	str=parser.FindSetting(L"User_font");
	skin.User_font=LoadSkinFont(str,L"Segoe UI",FW_NORMAL,18,false);
	str=parser.FindSetting(L"User_text_color");
	if (str)
		LoadSkinNumbers(str,(int*)&skin.User_text_color,1,true);
	else
		skin.User_text_color=0xFFFFFF;

	str=parser.FindSetting(L"User_glow_color");
	if (str)
		LoadSkinNumbers(str,(int*)&skin.User_glow_color,1,true);
	else
		skin.User_glow_color=0xFFFFFF;

	str=parser.FindSetting(L"User_glow_size");
	if (str)
		skin.User_glow_size=_wtol(str);
	else
		skin.User_glow_size=0;


	// SUB-MENU SECTION - describes the menu portion of the sub-menu
	str=parser.FindSetting(L"Submenu_background");
	if (str)
		LoadSkinNumbers(str,(int*)&skin.Submenu_background,1,true);
	else
		skin.Submenu_background=GetSysColor(COLOR_MENU);

	str=parser.FindSetting(L"Submenu_bitmap");
	if (str && (flags&LOADMENU_RESOURCES))
	{
		int id=_wtol(str);
		if (id)
		{
			str=parser.FindSetting(L"Submenu_bitmap_mask");
			int id2=str?_wtol(str):0;
			skin.Submenu_bitmap=LoadSkinBitmap(hMod,id,id2,skin.Submenu_bitmap32,skin.Submenu_background);
			if (!skin.Submenu_bitmap) return false;
			if (bRTL)
				MirrorBitmap(skin.Submenu_bitmap);
		}
	}
	memset(skin.Submenu_bitmap_slices_X,0,sizeof(skin.Submenu_bitmap_slices_X));
	str=parser.FindSetting(L"Submenu_bitmap_slices_X");
	if (str)
		LoadSkinNumbers(str,skin.Submenu_bitmap_slices_X+3,_countof(skin.Submenu_bitmap_slices_X)-3,false);
	str=parser.FindSetting(L"Submenu_bitmap_slices_Y");
	if (str)
		LoadSkinNumbers(str,skin.Submenu_bitmap_slices_Y,_countof(skin.Submenu_bitmap_slices_Y),false);
	else
		memset(skin.Submenu_bitmap_slices_Y,0,sizeof(skin.Submenu_bitmap_slices_Y));

	str=parser.FindSetting(L"Submenu_opacity");
	skin.Submenu_opacity=MenuSkin::OPACITY_SOLID;
	if (str && skin.Submenu_bitmap)
	{
		if (_wcsicmp(str,L"region")==0) skin.Submenu_opacity=MenuSkin::OPACITY_REGION;
		if (_wcsicmp(str,L"alpha")==0) skin.Submenu_opacity=MenuSkin::OPACITY_ALPHA;
		if (_wcsicmp(str,L"glass")==0) skin.Submenu_opacity=MenuSkin::OPACITY_GLASS;
		if (_wcsicmp(str,L"fullalpha")==0) skin.Submenu_opacity=MenuSkin::OPACITY_FULLALPHA;
		if (_wcsicmp(str,L"fullglass")==0) skin.Submenu_opacity=MenuSkin::OPACITY_FULLGLASS;
	}

	str=parser.FindSetting(L"Submenu_font");
	skin.Submenu_font=LoadSkinFont(str,NULL,0,0,true);

	str=parser.FindSetting(L"Submenu_glow_size");
	if (str)
		skin.Submenu_glow_size=_wtol(str);
	else
		skin.Submenu_glow_size=0;

	str=parser.FindSetting(L"Submenu_text_color");
	if (str)
		LoadSkinNumbers(str,(int*)skin.Submenu_text_color,_countof(skin.Submenu_text_color),true);
	else
	{
		skin.Submenu_text_color[0]=GetSysColor(COLOR_MENUTEXT);
		skin.Submenu_text_color[1]=GetSysColor(COLOR_HIGHLIGHTTEXT);
		skin.Submenu_text_color[2]=GetSysColor(COLOR_GRAYTEXT);
		skin.Submenu_text_color[3]=GetSysColor(COLOR_HIGHLIGHTTEXT);
	}
	str=parser.FindSetting(L"Submenu_arrow_color");
	if (str)
		LoadSkinNumbers(str,(int*)skin.Submenu_arrow_color,_countof(skin.Submenu_arrow_color),true);
	else
	{
		skin.Submenu_arrow_color[0]=skin.Submenu_text_color[0];
		skin.Submenu_arrow_color[1]=skin.Submenu_text_color[1];
	}

	str=parser.FindSetting(L"Submenu_padding");
	if (str)
		LoadSkinNumbers(str,(int*)&skin.Submenu_padding,4,false);
	else
		memset(&skin.Submenu_padding,0,sizeof(skin.Submenu_padding));
	str=parser.FindSetting(L"Submenu_offset");
	if (str)
		skin.Submenu_offset=_wtol(str);
	else
		skin.Submenu_offset=0;

	str=parser.FindSetting(L"AllPrograms_offset");
	if (str)
		skin.AllPrograms_offset=_wtol(str);
	else
		skin.AllPrograms_offset=0;

	str=parser.FindSetting(L"Submenu_selection");
	if (str)
	{
		if (str[0]=='#')
			LoadSkinNumbers(str,(int*)&skin.Submenu_selection.color,1,true);
		else
		{
			skin.Submenu_selectionColor=false;
			skin.Submenu_selection.bmp=NULL;
			if (flags&LOADMENU_RESOURCES)
			{
				int id=_wtol(str);
				if (id)
				{
					skin.Submenu_selection.bmp=LoadSkinBitmap(hMod,id,0,skin.Submenu_selection32,0);
					if (!skin.Submenu_selection.bmp) return false;
				}
			}

			str=parser.FindSetting(L"Submenu_selection_slices_X");
			if (str)
				LoadSkinNumbers(str,(int*)&skin.Submenu_selection_slices_X,_countof(skin.Submenu_selection_slices_X),false);
			else
				memset(skin.Submenu_selection_slices_X,0,sizeof(skin.Submenu_selection_slices_X));
		}
		str=parser.FindSetting(L"Submenu_selection_slices_Y");
		if (str)
			LoadSkinNumbers(str,(int*)&skin.Submenu_selection_slices_Y,_countof(skin.Submenu_selection_slices_Y),false);
		else
			memset(skin.Submenu_selection_slices_Y,0,sizeof(skin.Submenu_selection_slices_Y));
	}
	else
	{
		skin.Submenu_selection.color=GetSysColor(COLOR_MENUHILIGHT);
	}

	if (bRTL && !skin.Submenu_selectionColor && skin.Submenu_selection.bmp)
	{
		MirrorBitmap(skin.Submenu_selection.bmp);
		int q=skin.Submenu_selection_slices_X[0]; skin.Submenu_selection_slices_X[0]=skin.Submenu_selection_slices_X[2]; skin.Submenu_selection_slices_X[2]=q;
	}

	skin.Submenu_arrow_Size.cx=DEFAULT_ARROW_SIZE;
	skin.Submenu_arrow_Size.cy=0;
	str=parser.FindSetting(L"Submenu_arrow");
	if (str && (flags&LOADMENU_RESOURCES))
	{
		int id=_wtol(str);
		if (id)
		{
			skin.Submenu_arrow=LoadSkinBitmap(hMod,id,0,skin.Submenu_arrow32,0);
			if (!skin.Submenu_arrow) return false;
			if (bRTL)
				MirrorBitmap(skin.Submenu_arrow);
			BITMAP info;
			GetObject(skin.Submenu_arrow,sizeof(info),&info);
			skin.Submenu_arrow_Size.cx=info.bmWidth;
			skin.Submenu_arrow_Size.cy=info.bmHeight/2;
		}
	}

	str=parser.FindSetting(L"Submenu_arrow_padding");
	if (str)
		LoadSkinNumbers(str,(int*)&skin.Submenu_arrow_padding,2,false);
	else
		skin.Submenu_arrow_padding=DEFAULT_ARROW_PADDING;

	str=parser.FindSetting(L"Submenu_thin_frame");
	skin.Submenu_thin_frame=(str && _wtol(str));
	str=parser.FindSetting(L"Submenu_separator");
	if (str && (flags&LOADMENU_RESOURCES))
	{
		int id=_wtol(str);
		if (id)
		{
			skin.Submenu_separator=LoadSkinBitmap(hMod,id,0,skin.Submenu_separator32,0);
			if (!skin.Submenu_separator) return false;
			BITMAP info;
			GetObject(skin.Submenu_separator,sizeof(info),&info);
			skin.Submenu_separatorHeight=info.bmHeight;
		}
	}
	str=parser.FindSetting(L"Submenu_separator_slices_X");
	if (str)
		LoadSkinNumbers(str,skin.Submenu_separator_slices_X,_countof(skin.Submenu_separator_slices_X),false);
	else
		memset(skin.Submenu_separator_slices_X,0,sizeof(skin.Submenu_separator_slices_X));
	str=parser.FindSetting(L"Submenu_separatorV");
	skin.Submenu_separatorWidth=0;
	if (str && (flags&LOADMENU_RESOURCES))
	{
		int id=_wtol(str);
		if (id)
		{
			skin.Submenu_separatorV=LoadSkinBitmap(hMod,id,0,skin.Submenu_separatorV32,0);
			if (!skin.Submenu_separatorV) return false;
			BITMAP info;
			GetObject(skin.Submenu_separatorV,sizeof(info),&info);
			skin.Submenu_separatorWidth=info.bmWidth;
			if (bRTL)
				MirrorBitmap(skin.Submenu_separatorV);
		}
	}
	str=parser.FindSetting(L"Submenu_separator_slices_Y");
	if (str)
		LoadSkinNumbers(str,skin.Submenu_separator_slices_Y,_countof(skin.Submenu_separator_slices_Y),false);
	else
		memset(skin.Submenu_separator_slices_Y,0,sizeof(skin.Submenu_separator_slices_Y));

	if (skin.Main_bitmap_slices_X[1]==0)
	{
		skin.Main_bitmap_slices_X[0]=skin.Main_bitmap_slices_X[2]=0;
		memset(&skin.Caption_padding,0,sizeof(skin.Caption_padding));
	}

	str=parser.FindSetting(L"Submenu_icon_padding");
	if (str)
		LoadSkinNumbers(str,(int*)&skin.Submenu_icon_padding,4,false);
	else
		skin.Submenu_icon_padding=DEFAULT_ICON_PADDING;
	str=parser.FindSetting(L"Submenu_text_padding");
	if (str)
		LoadSkinNumbers(str,(int*)&skin.Submenu_text_padding,4,false);
	else
		skin.Submenu_text_padding=DEFAULT_TEXT_PADDING;

	str=parser.FindSetting(L"Submenu_pager");
	if (str && (flags&LOADMENU_RESOURCES))
	{
		int id=_wtol(str);
		if (id)
			skin.Submenu_pager=LoadSkinBitmap(hMod,id,0,skin.Submenu_pager32,0);
	}
	str=parser.FindSetting(L"Submenu_pager_slices_X");
	if (str)
		LoadSkinNumbers(str,skin.Submenu_pager_slices_X,_countof(skin.Submenu_pager_slices_X),false);
	else
		memset(skin.Submenu_pager_slices_X,0,sizeof(skin.Submenu_pager_slices_X));
	str=parser.FindSetting(L"Submenu_pager_slices_Y");
	if (str)
		LoadSkinNumbers(str,skin.Submenu_pager_slices_Y,_countof(skin.Submenu_pager_slices_Y),false);
	else
		memset(skin.Submenu_pager_slices_Y,0,sizeof(skin.Submenu_pager_slices_Y));

	if (bRTL && skin.Submenu_pager)
	{
		MirrorBitmap(skin.Submenu_pager);
		int q=skin.Submenu_pager_slices_X[0]; skin.Submenu_pager_slices_X[0]=skin.Submenu_pager_slices_X[2]; skin.Submenu_pager_slices_X[2]=q;
	}

	str=parser.FindSetting(L"Submenu_pager_arrows");
	if (str && (flags&LOADMENU_RESOURCES))
	{
		int id=_wtol(str);
		if (id)
		{
			skin.Submenu_pager_arrows=LoadSkinBitmap(hMod,id,0,skin.Submenu_pager_arrows32,0);
			BITMAP info;
			GetObject(skin.Submenu_pager_arrows,sizeof(info),&info);
			skin.Submenu_pager_arrow_Size.cx=info.bmWidth/2;
			skin.Submenu_pager_arrow_Size.cy=info.bmHeight/2;
		}
	}
	if (bRTL && skin.Submenu_pager_arrows)
		MirrorBitmap(skin.Submenu_pager_arrows);

	str=parser.FindSetting(L"Search_bitmap");
	if (str && (flags&LOADMENU_RESOURCES))
	{
		int id=_wtol(str);
		if (id)
		{
			skin.Search_bitmap=LoadSkinBitmap(hMod,id,0,skin.Search_bitmap32,0);
			if (!skin.Search_bitmap) return false;
		}
	}

	return true;
}

bool LoadMenuSkin( const wchar_t *fname, MenuSkin &skin, const wchar_t *variation, const wchar_t *optionsStr, unsigned int flags )
{
	MenuSkin::s_SkinError[0]=0;
	wchar_t path[_MAX_PATH];
	GetSkinsPath(path);

	if (!*fname || _wcsicmp(fname,LoadStringEx(IDS_DEFAULT_SKIN))==0)
	{
		LoadDefaultMenuSkin(skin,flags);
		return true;
	}

	std::vector<unsigned int> options;
	for (const wchar_t *str=optionsStr;*str;)
	{
		wchar_t token[256];
		str=GetToken(str,token,_countof(token),L"|");
		wchar_t *q;
		unsigned int hash=wcstoul(token,&q,16);
		options.push_back(hash);
	}

	if (_wcsicmp(fname,L"Custom")==0)
	{
		if (!LoadSkin(NULL,skin,variation,options,flags))
		{
			FreeMenuSkin(skin);
			return false;
		}
	}
	else
	{
		Strcat(path,_countof(path),fname);
		Strcat(path,_countof(path),L".skin");
		HMODULE hMod=LoadLibraryEx(path,NULL,LOAD_LIBRARY_AS_DATAFILE|LOAD_LIBRARY_AS_IMAGE_RESOURCE);
		if (!hMod)
		{
			wchar_t err[1024];
			GetErrorMessage(err,_countof(err),GetLastError());
			Sprintf(MenuSkin::s_SkinError,_countof(MenuSkin::s_SkinError),LoadStringEx(IDS_SKIN_ERR_LOAD),path,err);
			return false;
		}

		if (!LoadSkin(hMod,skin,variation,options,flags))
		{
			FreeMenuSkin(skin);
			FreeLibrary(hMod);
			return false;
		}

		FreeLibrary(hMod);
	}
	if (skin.version>MAX_SKIN_VERSION)
	{
		Strcpy(MenuSkin::s_SkinError,_countof(MenuSkin::s_SkinError),LoadStringEx(IDS_SKIN_ERR_VERSION));
		return false;
	}
	return true;
}

void LoadDefaultMenuSkin( MenuSkin &skin, unsigned int flags )
{
	LoadSkin(g_Instance,skin,NULL,std::vector<unsigned int>(),flags);
}

void FreeMenuSkin( MenuSkin &skin )
{
	skin.Reset();
}

void GetSkinsPath( wchar_t *path )
{
	GetModuleFileName(g_Instance,path,_MAX_PATH);
	*PathFindFileName(path)=0;
#ifdef BUILD_SETUP
	Strcat(path,_MAX_PATH,L"Skins\\");
#else
	Strcat(path,_MAX_PATH,L"..\\Skins\\");
#endif
}
