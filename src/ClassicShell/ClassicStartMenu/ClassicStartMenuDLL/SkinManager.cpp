// Classic Shell (c) 2009-2010, Ivo Beltchev
// The sources for Classic Shell are distributed under the MIT open source license

#include "stdafx.h"
#include "SkinManager.h"
#include "IconManager.h"
#include "ParseSettings.h"
#include "TranslationSettings.h"
#include "FNVHash.h"
#include <dwmapi.h>

const RECT DEFAULT_ICON_PADDING={3,3,3,3};
const RECT DEFAULT_TEXT_PADDING={1,0,4,0};

MenuSkin::MenuSkin( void )
{
	AboutIcon=NULL;
	Main_bitmap=NULL;
	Submenu_bitmap=NULL;
	Caption_font=NULL;
	Main_font=NULL;
	Main_selectionColor=true;
	Main_separator=NULL;
	Main_arrow=NULL;
	Submenu_font=NULL;
	Submenu_selectionColor=true;
	Submenu_separator=NULL;
	Submenu_arrow=NULL;
	Submenu_separatorV=NULL;
	Main_pager=NULL;
	Main_pager_arrows=NULL;
	Submenu_pager=NULL;
	Submenu_pager_arrows=NULL;
}

MenuSkin::~MenuSkin( void )
{
	Reset();
}

void MenuSkin::Reset( void )
{
	if (AboutIcon) DestroyIcon(AboutIcon);
	if (Main_bitmap) DeleteObject(Main_bitmap);
	if (Submenu_bitmap) DeleteObject(Submenu_bitmap);
	if (Caption_font) DeleteObject(Caption_font);
	if (Main_font) DeleteObject(Main_font);
	if (!Main_selectionColor && Main_selection.bmp) DeleteObject(Main_selection.bmp);
	if (Main_separator) DeleteObject(Main_separator);
	if (Main_arrow) DeleteObject(Main_arrow);
	if (Submenu_font) DeleteObject(Submenu_font);
	if (!Submenu_selectionColor && Submenu_selection.bmp) DeleteObject(Submenu_selection.bmp);
	if (Submenu_separator) DeleteObject(Submenu_separator);
	if (Submenu_separatorV) DeleteObject(Submenu_separatorV);
	if (Submenu_arrow) DeleteObject(Submenu_arrow);
	if (Main_pager) DeleteObject(Main_pager);
	if (Main_pager_arrows) DeleteObject(Main_pager_arrows);
	if (Submenu_pager) DeleteObject(Submenu_pager);
	if (Submenu_pager_arrows) DeleteObject(Submenu_pager_arrows);

	AboutIcon=NULL;
	Main_bitmap=NULL;
	Submenu_bitmap=NULL;
	Caption_font=NULL;
	Main_font=NULL;
	Main_selectionColor=true;
	Main_separator=NULL;
	Main_arrow=NULL;
	Submenu_font=NULL;
	Submenu_selectionColor=true;
	Submenu_separator=NULL;
	Submenu_separatorV=NULL;
	Submenu_arrow=NULL;
	Main_pager=NULL;
	Main_pager_arrows=NULL;
	Submenu_pager=NULL;
	Submenu_pager_arrows=NULL;
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

static HFONT LoadSkinFont( const wchar_t *str, const wchar_t *name, int weight, int size )
{
	wchar_t token[256];
	if (str)
	{
		str=GetToken(str,token,_countof(token),L", \t");
		name=token;
		wchar_t token2[256];
		str=GetToken(str,token2,_countof(token2),L", \t");
		if (_wcsicmp(token2,L"bold")==0)
			weight=FW_BOLD;
		else
			weight=FW_NORMAL;
		str=GetToken(str,token2,_countof(token2),L", \t");
		size=_wtol(token2);
	}
	else if (!name)
	{
		// get the default menu font
		NONCLIENTMETRICS metrics={sizeof(metrics)};
		SystemParametersInfo(SPI_GETNONCLIENTMETRICS,NULL,&metrics,0);
		return CreateFontIndirect(&metrics.lfMenuFont);
	}
	return CreateFont(size*CIconManager::GetDPI()/72,0,0,0,weight,0,0,0,DEFAULT_CHARSET,OUT_DEFAULT_PRECIS,CLIP_DEFAULT_PRECIS,CLEARTYPE_QUALITY,DEFAULT_PITCH,name);
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
	HBITMAP bmp;
	if (hMod)
		bmp=(HBITMAP)LoadImage(hMod,MAKEINTRESOURCE(index),IMAGE_BITMAP,0,0,LR_CREATEDIBSECTION);
	else
	{
		wchar_t path[_MAX_PATH];
		GetSkinsPath(path);
		wchar_t fname[_MAX_PATH];
		Sprintf(fname,_countof(fname),L"%s%d.bmp",path,index);
		bmp=(HBITMAP)LoadImage(NULL,fname,IMAGE_BITMAP,0,0,LR_CREATEDIBSECTION|LR_LOADFROMFILE);
	}
	if (!bmp) return NULL;

	HBITMAP bmpMask=NULL;
	BITMAP infoMask;
	if (maskIndex>0)
	{
		if (hMod)
			bmpMask=(HBITMAP)LoadImage(hMod,MAKEINTRESOURCE(maskIndex),IMAGE_BITMAP,0,0,LR_CREATEDIBSECTION);
		else
		{
			wchar_t path[_MAX_PATH];
			GetSkinsPath(path);
			wchar_t fname[_MAX_PATH];
			Sprintf(fname,_countof(fname),L"%s%d.bmp",path,maskIndex);
			bmpMask=(HBITMAP)LoadImage(NULL,fname,IMAGE_BITMAP,0,0,LR_CREATEDIBSECTION|LR_LOADFROMFILE);
		}
		if (bmpMask)
			GetObject(bmpMask,sizeof(infoMask),&infoMask);
	}

	BITMAP info;
	GetObject(bmp,sizeof(info),&info);

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

// Load the skin from the module. If hMod is NULL loads the "custom" skin from 1.txt
static bool LoadSkin( HMODULE hMod, MenuSkin &skin, const wchar_t *variation, const std::vector<unsigned int> &options, bool bNoResources )
{
	skin.version=1;
	CSkinParser parser;
	if (hMod)
	{
		HRSRC hResInfo=FindResource(hMod,MAKEINTRESOURCE(1),L"SKIN");
		if (!hResInfo) return false;
		if (!parser.LoadText(hMod,hResInfo)) return false;
	}
	else
	{
		wchar_t path[_MAX_PATH];
		GetSkinsPath(path);
		Strcat(path,_countof(path),L"1.txt");
		if (!parser.LoadText(path)) return false;
	}
	parser.ParseText();

	const wchar_t *str;

	for (int i=1;;i++)
	{
		char name[20];
		Sprintf(name,_countof(name),"Variation%d",i);
		str=parser.FindSetting(name);
		if (str)
		{
			wchar_t token[256];
			str=GetToken(str,token,_countof(token),L", \t");
			int  res=_wtol(token);
			str=GetToken(str,token,_countof(token),L", \t");
			skin.Variations.push_back(std::pair<int,CString>(res,token));
		}
		else
			break;
	}

	if (variation)
	{
		for (std::vector<std::pair<int,CString>>::const_iterator it=skin.Variations.begin();it!=skin.Variations.end();++it)
			if (wcscmp(variation,it->second)==0)
			{
				if (hMod)
				{
					HRSRC hResInfo=FindResource(hMod,MAKEINTRESOURCE(it->first),L"SKIN");
					if (!hResInfo) break;
					if (!parser.LoadVariation(hMod,hResInfo)) break;
				}
				else
				{
					wchar_t path[_MAX_PATH];
					GetSkinsPath(path);
					wchar_t name[20];
					Sprintf(name,_countof(name),L"%d.txt",it->first);
					Strcat(path,_countof(path),name);
					if (!parser.LoadVariation(path)) break;
				}

				break;
			}
	}

	// find options
	std::vector<const wchar_t*> values; // list of true values
	for (int i=0;;i++)
	{
		MenuSkin::Option option;
		if (!parser.ParseOption(option.name,option.label,option.value,i))
			break;
		if (!option.name.IsEmpty())
		{
			skin.Options.push_back(option);
			unsigned int hash=CalcFNVHash(option.name)&0xFFFFFFFE;
			bool bValue=option.value;
			for (std::vector<unsigned int>::const_iterator it=options.begin();it!=options.end();++it)
				if ((*it&0xFFFFFFFE)==hash)
				{
					bValue=(*it&1)!=0;
					break;
				}
			if (bValue)
				values.push_back(option.name);
		}
	}
	if (values.empty())
		parser.FilterConditions(NULL,0);
	else
		parser.FilterConditions(&values[0],(int)values.size());

	// parse settings
	str=parser.FindSetting("About");
	if (str)
	{
		skin.About=str;
		skin.About.Replace(L"\\n",L"\n");
	}
	else skin.About.Empty();

	str=parser.FindSetting("AboutIcon");
	if (str) skin.AboutIcon=LoadSkinIcon(hMod,_wtol(str));

	str=parser.FindSetting("Version");
	if (str)
		skin.version=_wtol(str);

	skin.ForceRTL=false;
	if (!hMod)
	{
		str=parser.FindSetting("ForceRTL");
		if (str && _wtol(str))
			skin.ForceRTL=true;
	}

	bool bRTL=skin.ForceRTL || IsLanguageRTL();

	// CAPTION SECTION - describes the caption portion of the main menu
	str=parser.FindSetting("Caption_font");
	skin.Caption_font=LoadSkinFont(str,L"Segoe UI",FW_NORMAL,18);
	str=parser.FindSetting("Caption_text_color");
	if (str)
		LoadSkinNumbers(str,(int*)&skin.Caption_text_color,1,true);
	else
		skin.Caption_text_color=0xFFFFFF;

	str=parser.FindSetting("Caption_glow_color");
	if (str)
		LoadSkinNumbers(str,(int*)&skin.Caption_glow_color,1,true);
	else
		skin.Caption_glow_color=0xFFFFFF;

	str=parser.FindSetting("Caption_glow_size");
	if (str)
		skin.Caption_glow_size=_wtol(str);
	else
		skin.Caption_glow_size=0;

	str=parser.FindSetting("Caption_padding");
	if (str)
		LoadSkinNumbers(str,(int*)&skin.Caption_padding,4,false);
	else
		memset(&skin.Caption_padding,0,sizeof(skin.Caption_padding));


	// MENU SECTION - describes the menu portion of the main menu
	str=parser.FindSetting("Main_background");
	if (str)
		LoadSkinNumbers(str,(int*)&skin.Main_background,1,true);
	else
		skin.Main_background=GetSysColor(COLOR_MENU);

	str=parser.FindSetting("Main_bitmap");
	if (str && !bNoResources)
	{
		int id=_wtol(str);
		if (id)
		{
			str=parser.FindSetting("Main_bitmap_mask");
			int id2=str?_wtol(str):0;
			skin.Main_bitmap=LoadSkinBitmap(hMod,id,id2,skin.Main_bitmap32,skin.Main_background);
			if (!skin.Main_bitmap) return false;
		}
	}
	str=parser.FindSetting("Main_bitmap_slices_X");
	if (str)
		LoadSkinNumbers(str,skin.Main_bitmap_slices_X,_countof(skin.Main_bitmap_slices_X),false);
	else
		memset(skin.Main_bitmap_slices_X,0,sizeof(skin.Main_bitmap_slices_X));
	str=parser.FindSetting("Main_bitmap_slices_Y");
	if (str)
		LoadSkinNumbers(str,skin.Main_bitmap_slices_Y,_countof(skin.Main_bitmap_slices_Y),false);
	else
		memset(skin.Main_bitmap_slices_Y,0,sizeof(skin.Main_bitmap_slices_Y));
	str=parser.FindSetting("Main_opacity");
	skin.Main_opacity=MenuSkin::OPACITY_SOLID;
	if (str && skin.Main_bitmap)
	{
		if (_wcsicmp(str,L"region")==0) skin.Main_opacity=MenuSkin::OPACITY_REGION;
		if (_wcsicmp(str,L"alpha")==0) skin.Main_opacity=MenuSkin::OPACITY_ALPHA;
		if (_wcsicmp(str,L"glass")==0) skin.Main_opacity=MenuSkin::OPACITY_GLASS;
		if (_wcsicmp(str,L"fullalpha")==0) skin.Main_opacity=MenuSkin::OPACITY_FULLALPHA;
		if (_wcsicmp(str,L"fullglass")==0) skin.Main_opacity=MenuSkin::OPACITY_FULLGLASS;
	}
	str=parser.FindSetting("Main_large_icons");
	skin.Main_large_icons=str && _wtol(str);

	str=parser.FindSetting("Main_font");
	skin.Main_font=LoadSkinFont(str,NULL,0,0);

	str=parser.FindSetting("Main_glow_size");
	if (str)
		skin.Main_glow_size=_wtol(str);
	else
		skin.Main_glow_size=0;

	str=parser.FindSetting("Main_text_color");
	if (str)
		LoadSkinNumbers(str,(int*)skin.Main_text_color,_countof(skin.Main_text_color),true);
	else
	{
		skin.Main_text_color[0]=GetSysColor(COLOR_MENUTEXT);
		skin.Main_text_color[1]=GetSysColor(COLOR_HIGHLIGHTTEXT);
		skin.Main_text_color[2]=GetSysColor(COLOR_GRAYTEXT);
		skin.Main_text_color[3]=GetSysColor(COLOR_HIGHLIGHTTEXT);
	}
	str=parser.FindSetting("Main_arrow_color");
	if (str)
		LoadSkinNumbers(str,(int*)skin.Main_arrow_color,_countof(skin.Main_arrow_color),true);
	else
	{
		skin.Main_arrow_color[0]=skin.Main_text_color[0];
		skin.Main_arrow_color[1]=skin.Main_text_color[1];
	}

	str=parser.FindSetting("Main_padding");
	if (str)
		LoadSkinNumbers(str,(int*)&skin.Main_padding,4,false);
	else
		memset(&skin.Main_padding,0,sizeof(skin.Main_padding));

	str=parser.FindSetting("Main_selection");
	if (str)
	{
		if (str[0]=='#')
			LoadSkinNumbers(str,(int*)&skin.Main_selection.color,1,true);
		else
		{
			skin.Main_selectionColor=false;
			if (!bNoResources)
			{
				int id=_wtol(str);
				if (id)
				{
					skin.Main_selection.bmp=LoadSkinBitmap(hMod,id,0,skin.Main_selection32,0);
					if (!skin.Main_selection.bmp) return false;
				}
			}

			str=parser.FindSetting("Main_selection_slices_X");
			if (str)
				LoadSkinNumbers(str,(int*)&skin.Main_selection_slices_X,_countof(skin.Main_selection_slices_X),false);
			else
				memset(skin.Main_selection_slices_X,0,sizeof(skin.Main_selection_slices_X));
		}
		str=parser.FindSetting("Main_selection_slices_Y");
		if (str)
			LoadSkinNumbers(str,(int*)&skin.Main_selection_slices_Y,_countof(skin.Main_selection_slices_Y),false);
		else
			memset(skin.Main_selection_slices_Y,0,sizeof(skin.Main_selection_slices_Y));
	}
	else
	{
		skin.Main_selection.color=GetSysColor(COLOR_MENUHILIGHT);
	}

	if (bRTL && !skin.Main_selectionColor && skin.Main_selection.bmp)
	{
		MirrorBitmap(skin.Main_selection.bmp);
		int q=skin.Main_selection_slices_X[0]; skin.Main_selection_slices_X[0]=skin.Main_selection_slices_X[2]; skin.Main_selection_slices_X[2]=q;
	}

	str=parser.FindSetting("Main_arrow");
	if (str && !bNoResources)
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

	str=parser.FindSetting("Main_thin_frame");
	skin.Main_thin_frame=(str && _wtol(str));
	str=parser.FindSetting("Main_separator");
	if (str && !bNoResources)
	{
		int id=_wtol(str);
		if (id)
		{
			skin.Main_separator=LoadSkinBitmap(hMod,id,0,skin.Main_separator32,0);
			if (!skin.Main_separator) return false;
			BITMAP info;
			GetObject(skin.Main_separator,sizeof(info),&info);
			skin.Main_separatorHeight=info.bmHeight;
		}
	}
	str=parser.FindSetting("Main_separator_slices_X");
	if (str)
	{
		LoadSkinNumbers(str,skin.Main_separator_slices_X,_countof(skin.Main_separator_slices_X),false);
	}
	else
		memset(skin.Main_separator_slices_X,0,sizeof(skin.Main_separator_slices_X));

	str=parser.FindSetting("Main_icon_padding");
	if (str)
		LoadSkinNumbers(str,(int*)&skin.Main_icon_padding,4,false);
	else
		skin.Main_icon_padding=DEFAULT_ICON_PADDING;
	str=parser.FindSetting("Main_text_padding");
	if (str)
		LoadSkinNumbers(str,(int*)&skin.Main_text_padding,4,false);
	else
		skin.Main_text_padding=DEFAULT_TEXT_PADDING;

	str=parser.FindSetting("Main_pager");
	if (str && !bNoResources)
	{
		int id=_wtol(str);
		if (id)
			skin.Main_pager=LoadSkinBitmap(hMod,id,0,skin.Main_pager32,0);
	}
	str=parser.FindSetting("Main_pager_slices_X");
	if (str)
		LoadSkinNumbers(str,skin.Main_pager_slices_X,_countof(skin.Main_pager_slices_X),false);
	else
		memset(skin.Main_pager_slices_X,0,sizeof(skin.Main_pager_slices_X));
	str=parser.FindSetting("Main_pager_slices_Y");
	if (str)
		LoadSkinNumbers(str,skin.Main_pager_slices_Y,_countof(skin.Main_pager_slices_Y),false);
	else
		memset(skin.Main_pager_slices_Y,0,sizeof(skin.Main_pager_slices_Y));
	if (bRTL && skin.Main_pager)
	{
		MirrorBitmap(skin.Main_pager);
		int q=skin.Main_pager_slices_X[0]; skin.Main_pager_slices_X[0]=skin.Main_pager_slices_X[2]; skin.Main_pager_slices_X[2]=q;
	}

	str=parser.FindSetting("Main_pager_arrows");
	if (str && !bNoResources)
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

	// SUB-MENU SECTION - describes the menu portion of the sub-menu
	str=parser.FindSetting("Submenu_background");
	if (str)
		LoadSkinNumbers(str,(int*)&skin.Submenu_background,1,true);
	else
		skin.Submenu_background=GetSysColor(COLOR_MENU);

	str=parser.FindSetting("Submenu_bitmap");
	if (str && !bNoResources)
	{
		int id=_wtol(str);
		if (id)
		{
			str=parser.FindSetting("Submenu_bitmap_mask");
			int id2=str?_wtol(str):0;
			skin.Submenu_bitmap=LoadSkinBitmap(hMod,id,id2,skin.Submenu_bitmap32,skin.Submenu_background);
			if (!skin.Submenu_bitmap) return false;
			if (bRTL)
				MirrorBitmap(skin.Submenu_bitmap);
		}
	}
	memset(skin.Submenu_bitmap_slices_X,0,sizeof(skin.Submenu_bitmap_slices_X));
	str=parser.FindSetting("Submenu_bitmap_slices_X");
	if (str)
		LoadSkinNumbers(str,skin.Submenu_bitmap_slices_X+3,_countof(skin.Submenu_bitmap_slices_X)-3,false);
	str=parser.FindSetting("Submenu_bitmap_slices_Y");
	if (str)
		LoadSkinNumbers(str,skin.Submenu_bitmap_slices_Y,_countof(skin.Submenu_bitmap_slices_Y),false);
	else
		memset(skin.Submenu_bitmap_slices_Y,0,sizeof(skin.Submenu_bitmap_slices_Y));

	str=parser.FindSetting("Submenu_opacity");
	skin.Submenu_opacity=MenuSkin::OPACITY_SOLID;
	if (str && skin.Submenu_bitmap)
	{
		if (_wcsicmp(str,L"region")==0) skin.Submenu_opacity=MenuSkin::OPACITY_REGION;
		if (_wcsicmp(str,L"alpha")==0) skin.Submenu_opacity=MenuSkin::OPACITY_ALPHA;
		if (_wcsicmp(str,L"glass")==0) skin.Submenu_opacity=MenuSkin::OPACITY_GLASS;
		if (_wcsicmp(str,L"fullalpha")==0) skin.Submenu_opacity=MenuSkin::OPACITY_FULLALPHA;
		if (_wcsicmp(str,L"fullglass")==0) skin.Submenu_opacity=MenuSkin::OPACITY_FULLGLASS;
	}

	str=parser.FindSetting("Submenu_font");
	skin.Submenu_font=LoadSkinFont(str,NULL,0,0);

	str=parser.FindSetting("Submenu_glow_size");
	if (str)
		skin.Submenu_glow_size=_wtol(str);
	else
		skin.Submenu_glow_size=0;

	str=parser.FindSetting("Submenu_text_color");
	if (str)
		LoadSkinNumbers(str,(int*)skin.Submenu_text_color,_countof(skin.Submenu_text_color),true);
	else
	{
		skin.Submenu_text_color[0]=GetSysColor(COLOR_MENUTEXT);
		skin.Submenu_text_color[1]=GetSysColor(COLOR_HIGHLIGHTTEXT);
		skin.Submenu_text_color[2]=GetSysColor(COLOR_GRAYTEXT);
		skin.Submenu_text_color[3]=GetSysColor(COLOR_HIGHLIGHTTEXT);
	}
	str=parser.FindSetting("Submenu_arrow_color");
	if (str)
		LoadSkinNumbers(str,(int*)skin.Submenu_arrow_color,_countof(skin.Submenu_arrow_color),true);
	else
	{
		skin.Submenu_arrow_color[0]=skin.Submenu_text_color[0];
		skin.Submenu_arrow_color[1]=skin.Submenu_text_color[1];
	}

	str=parser.FindSetting("Submenu_padding");
	if (str)
		LoadSkinNumbers(str,(int*)&skin.Submenu_padding,4,false);
	else
		memset(&skin.Submenu_padding,0,sizeof(skin.Submenu_padding));
	str=parser.FindSetting("Submenu_offset");
	if (str)
		skin.Submenu_offset=_wtol(str);
	else
		skin.Submenu_offset=0;

	str=parser.FindSetting("Submenu_selection");
	if (str)
	{
		if (str[0]=='#')
			LoadSkinNumbers(str,(int*)&skin.Submenu_selection.color,1,true);
		else
		{
			skin.Submenu_selectionColor=false;
			if (!bNoResources)
			{
				int id=_wtol(str);
				if (id)
				{
					skin.Submenu_selection.bmp=LoadSkinBitmap(hMod,id,0,skin.Submenu_selection32,0);
					if (!skin.Submenu_selection.bmp) return false;
				}
			}

			str=parser.FindSetting("Submenu_selection_slices_X");
			if (str)
				LoadSkinNumbers(str,(int*)&skin.Submenu_selection_slices_X,_countof(skin.Submenu_selection_slices_X),false);
			else
				memset(skin.Submenu_selection_slices_X,0,sizeof(skin.Submenu_selection_slices_X));
		}
		str=parser.FindSetting("Submenu_selection_slices_Y");
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

	str=parser.FindSetting("Submenu_arrow");
	if (str && !bNoResources)
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

	str=parser.FindSetting("Submenu_thin_frame");
	skin.Submenu_thin_frame=(str && _wtol(str));
	str=parser.FindSetting("Submenu_separator");
	if (str && !bNoResources)
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
	str=parser.FindSetting("Submenu_separator_slices_X");
	if (str)
		LoadSkinNumbers(str,skin.Submenu_separator_slices_X,_countof(skin.Submenu_separator_slices_X),false);
	else
		memset(skin.Submenu_separator_slices_X,0,sizeof(skin.Submenu_separator_slices_X));
	str=parser.FindSetting("Submenu_separatorV");
	skin.Submenu_separatorWidth=0;
	if (str && !bNoResources)
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
	str=parser.FindSetting("Submenu_separator_slices_Y");
	if (str)
		LoadSkinNumbers(str,skin.Submenu_separator_slices_Y,_countof(skin.Submenu_separator_slices_Y),false);
	else
		memset(skin.Submenu_separator_slices_Y,0,sizeof(skin.Submenu_separator_slices_Y));

	if (skin.Main_bitmap_slices_X[1]==0)
	{
		skin.Main_bitmap_slices_X[0]=skin.Main_bitmap_slices_X[2]=0;
		memset(&skin.Caption_padding,0,sizeof(skin.Caption_padding));
	}

	str=parser.FindSetting("Submenu_icon_padding");
	if (str)
		LoadSkinNumbers(str,(int*)&skin.Submenu_icon_padding,4,false);
	else
		skin.Submenu_icon_padding=DEFAULT_ICON_PADDING;
	str=parser.FindSetting("Submenu_text_padding");
	if (str)
		LoadSkinNumbers(str,(int*)&skin.Submenu_text_padding,4,false);
	else
		skin.Submenu_text_padding=DEFAULT_TEXT_PADDING;

	str=parser.FindSetting("Submenu_pager");
	if (str && !bNoResources)
	{
		int id=_wtol(str);
		if (id)
			skin.Submenu_pager=LoadSkinBitmap(hMod,id,0,skin.Submenu_pager32,0);
	}
	str=parser.FindSetting("Submenu_pager_slices_X");
	if (str)
		LoadSkinNumbers(str,skin.Submenu_pager_slices_X,_countof(skin.Submenu_pager_slices_X),false);
	else
		memset(skin.Submenu_pager_slices_X,0,sizeof(skin.Submenu_pager_slices_X));
	str=parser.FindSetting("Submenu_pager_slices_Y");
	if (str)
		LoadSkinNumbers(str,skin.Submenu_pager_slices_Y,_countof(skin.Submenu_pager_slices_Y),false);
	else
		memset(skin.Submenu_pager_slices_Y,0,sizeof(skin.Submenu_pager_slices_Y));

	if (bRTL && skin.Submenu_pager)
	{
		MirrorBitmap(skin.Submenu_pager);
		int q=skin.Submenu_pager_slices_X[0]; skin.Submenu_pager_slices_X[0]=skin.Submenu_pager_slices_X[2]; skin.Submenu_pager_slices_X[2]=q;
	}

	str=parser.FindSetting("Submenu_pager_arrows");
	if (str && !bNoResources)
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

	return true;
}

bool LoadMenuSkin( const wchar_t *fname, MenuSkin &skin, const wchar_t *variation, const std::vector<unsigned int> &options, bool bNoResources )
{
	wchar_t path[_MAX_PATH];
	GetSkinsPath(path);

	if (_wcsicmp(fname,L"<Default>")==0)
	{
		LoadDefaultMenuSkin(skin,bNoResources);
		return true;
	}
	if (_wcsicmp(fname,L"Custom")==0)
	{
		if (!LoadSkin(NULL,skin,variation,options,bNoResources))
		{
			FreeMenuSkin(skin);
			return false;
		}
	}
	else
	{
		Strcat(path,_countof(path),fname);
		Strcat(path,_countof(path),L".skin");
		HMODULE hMod=LoadLibraryEx(path,NULL,LOAD_LIBRARY_AS_DATAFILE);
		if (!hMod)
			return false;

		if (!LoadSkin(hMod,skin,variation,options,bNoResources))
		{
			FreeMenuSkin(skin);
			FreeLibrary(hMod);
			return false;
		}

		FreeLibrary(hMod);
	}
	return true;
}

void LoadDefaultMenuSkin( MenuSkin &skin, bool bNoResources )
{
	LoadSkin(g_Instance,skin,NULL,std::vector<unsigned int>(),bNoResources);
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
