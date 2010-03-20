// ## MenuContainer.h
// Classic Shell (c) 2009-2010, Ivo Beltchev
// The sources for Classic Shell are distributed under the MIT open source license

// MenuPaint.cpp - handles the painting functionality of CMenuContainer

#include "stdafx.h"
#include "IconManager.h"
#include "MenuContainer.h"
#include "ClassicStartMenuDLL.h"
#include "GlobalSettings.h"
#include <vsstyle.h>
#include <dwmapi.h>
#include <algorithm>

void CMenuContainer::MarginsBlit( HDC hSrc, HDC hDst, const RECT &rSrc, const RECT &rDst, const RECT &rMargins, bool bAlpha, bool bRtlOffset )
{
	int x0a=rDst.left;
	int x1a=rDst.left+rMargins.left;
	int x2a=rDst.right-rMargins.right;
	int x3a=rDst.right;
	int x0b=rSrc.left;
	int x1b=rSrc.left+rMargins.left;
	int x2b=rSrc.right-rMargins.right;
	int x3b=rSrc.right;

	int y0a=rDst.top;
	int y1a=rDst.top+rMargins.top;
	int y2a=rDst.bottom-rMargins.bottom;
	int y3a=rDst.bottom;
	int y0b=rSrc.top;
	int y1b=rSrc.top+rMargins.top;
	int y2b=rSrc.bottom-rMargins.bottom;
	int y3b=rSrc.bottom;

	SetStretchBltMode(hDst,COLORONCOLOR);
	if (bAlpha)
	{
		BLENDFUNCTION func={AC_SRC_OVER,0,255,AC_SRC_ALPHA};
		if (x0a<x1a && y0a<y1a && x0b<x1b && y0b<y1b) AlphaBlend(hDst,x0a,y0a,x1a-x0a,y1a-y0a,hSrc,x0b,y0b,x1b-x0b,y1b-y0b,func);
		if (x1a<x2a && y0a<y1a && x1b<x2b && y0b<y1b) AlphaBlend(hDst,x1a,y0a,x2a-x1a,y1a-y0a,hSrc,x1b,y0b,x2b-x1b,y1b-y0b,func);
		if (x2a<x3a && y0a<y1a && x2b<x3b && y0b<y1b) AlphaBlend(hDst,x2a,y0a,x3a-x2a,y1a-y0a,hSrc,x2b,y0b,x3b-x2b,y1b-y0b,func);

		if (x0a<x1a && y1a<y2a && x0b<x1b && y1b<y2b) AlphaBlend(hDst,x0a,y1a,x1a-x0a,y2a-y1a,hSrc,x0b,y1b,x1b-x0b,y2b-y1b,func);
		if (x1a<x2a && y1a<y2a && x1b<x2b && y1b<y2b) AlphaBlend(hDst,x1a,y1a,x2a-x1a,y2a-y1a,hSrc,x1b,y1b,x2b-x1b,y2b-y1b,func);
		if (x2a<x3a && y1a<y2a && x2b<x3b && y1b<y2b) AlphaBlend(hDst,x2a,y1a,x3a-x2a,y2a-y1a,hSrc,x2b,y1b,x3b-x2b,y2b-y1b,func);

		if (x0a<x1a && y2a<y3a && x0b<x1b && y2b<y3b) AlphaBlend(hDst,x0a,y2a,x1a-x0a,y3a-y2a,hSrc,x0b,y2b,x1b-x0b,y3b-y2b,func);
		if (x1a<x2a && y2a<y3a && x1b<x2b && y2b<y3b) AlphaBlend(hDst,x1a,y2a,x2a-x1a,y3a-y2a,hSrc,x1b,y2b,x2b-x1b,y3b-y2b,func);
		if (x2a<x3a && y2a<y3a && x2b<x3b && y2b<y3b) AlphaBlend(hDst,x2a,y2a,x3a-x2a,y3a-y2a,hSrc,x2b,y2b,x3b-x2b,y3b-y2b,func);
	}
	else
	{
		int o=bRtlOffset?1:0;
		if (x0a<x1a && y0a<y1a && x0b<x1b && y0b<y1b) StretchBlt(hDst,x0a-o,y0a,x1a-x0a+o,y1a-y0a,hSrc,x0b-o,y0b,x1b-x0b+o,y1b-y0b,SRCCOPY);
		if (x1a<x2a && y0a<y1a && x1b<x2b && y0b<y1b) StretchBlt(hDst,x1a,y0a,x2a-x1a,y1a-y0a,hSrc,x1b,y0b,x2b-x1b,y1b-y0b,SRCCOPY);
		if (x2a<x3a && y0a<y1a && x2b<x3b && y0b<y1b) StretchBlt(hDst,x2a-o,y0a,x3a-x2a+o,y1a-y0a,hSrc,x2b-o,y0b,x3b-x2b+o,y1b-y0b,SRCCOPY);

		if (x0a<x1a && y1a<y2a && x0b<x1b && y1b<y2b) StretchBlt(hDst,x0a,y1a,x1a-x0a,y2a-y1a,hSrc,x0b,y1b,x1b-x0b,y2b-y1b,SRCCOPY);
		if (x1a<x2a && y1a<y2a && x1b<x2b && y1b<y2b) StretchBlt(hDst,x1a,y1a,x2a-x1a,y2a-y1a,hSrc,x1b,y1b,x2b-x1b,y2b-y1b,SRCCOPY);
		if (x2a<x3a && y1a<y2a && x2b<x3b && y1b<y2b) StretchBlt(hDst,x2a,y1a,x3a-x2a,y2a-y1a,hSrc,x2b,y1b,x3b-x2b,y2b-y1b,SRCCOPY);

		if (x0a<x1a && y2a<y3a && x0b<x1b && y2b<y3b) StretchBlt(hDst,x0a-o,y2a,x1a-x0a+o,y3a-y2a,hSrc,x0b-o,y2b,x1b-x0b+o,y3b-y2b,SRCCOPY);
		if (x1a<x2a && y2a<y3a && x1b<x2b && y2b<y3b) StretchBlt(hDst,x1a,y2a,x2a-x1a,y3a-y2a,hSrc,x1b,y2b,x2b-x1b,y3b-y2b,SRCCOPY);
		if (x2a<x3a && y2a<y3a && x2b<x3b && y2b<y3b) StretchBlt(hDst,x2a-o,y2a,x3a-x2a+o,y3a-y2a,hSrc,x2b-o,y2b,x3b-x2b+o,y3b-y2b,SRCCOPY);
	}
}

// Creates the bitmap for the background
void CMenuContainer::CreateBackground( int width, int height )
{
	// get the text from the ini file or from the registry
	CRegKey regTitle;
	wchar_t title[256]=L"Windows";
	const wchar_t *setting=FindSetting("MenuCaption");
	if (setting)
		Strcpy(title,_countof(title),setting);
	else
	{
		ULONG size=_countof(title);
		if (regTitle.Open(HKEY_LOCAL_MACHINE,L"Software\\Microsoft\\Windows NT\\CurrentVersion",KEY_READ)==ERROR_SUCCESS)
			regTitle.QueryStringValue(L"ProductName",title,&size);
	}

	HBITMAP bmpSkin=m_pParent?s_Skin.Submenu_bitmap:s_Skin.Main_bitmap;
	bool b32=m_pParent?s_Skin.Submenu_bitmap32:s_Skin.Main_bitmap32;
	const int *slicesX=m_pParent?s_Skin.Submenu_bitmap_slices_X:s_Skin.Main_bitmap_slices_X;
	const int *slicesY=m_pParent?s_Skin.Submenu_bitmap_slices_Y:s_Skin.Main_bitmap_slices_Y;
	bool bCaption=(slicesX[1]>0);
	MenuSkin::TOpacity opacity=m_pParent?s_Skin.Submenu_opacity:s_Skin.Main_opacity;

	HDC hdcTemp=CreateCompatibleDC(NULL);

	HFONT font0=NULL;
	if (bCaption)
		font0=(HFONT)SelectObject(hdcTemp,s_Skin.Caption_font);

	RECT rc={0,0,0,0};
	DTTOPTS opts={sizeof(opts),DTT_COMPOSITED|DTT_CALCRECT};
	if (bCaption)
	{
		if (s_Theme)
			DrawThemeTextEx(s_Theme,hdcTemp,0,0,title,-1,DT_NOPREFIX|DT_SINGLELINE|DT_CALCRECT,&rc,&opts);
		else
			DrawText(hdcTemp,title,-1,&rc,DT_NOPREFIX|DT_SINGLELINE|DT_CALCRECT);
	}
	int textWidth=0, textHeight=0;
	if (!m_pParent)
	{
		textWidth=rc.right+s_Skin.Caption_padding.top+s_Skin.Caption_padding.bottom;
		textHeight=rc.bottom+s_Skin.Caption_padding.left+s_Skin.Caption_padding.right;
	}
	int total=slicesX[0]+slicesX[2];
	if (textHeight<total) textHeight=total;

	int totalWidth=textHeight+width;
	totalWidth+=m_pParent?(s_Skin.Submenu_padding.left+s_Skin.Submenu_padding.right):(s_Skin.Main_padding.left+s_Skin.Main_padding.right);
	total+=textHeight+slicesX[3]+slicesX[5];
	if (totalWidth<total) totalWidth=total;

	int totalHeight=height;
	totalHeight+=m_pParent?(s_Skin.Submenu_padding.top+s_Skin.Submenu_padding.bottom):(s_Skin.Main_padding.top+s_Skin.Main_padding.bottom);
	total=slicesY[0]+slicesY[2];
	if (totalHeight<total) totalHeight=total;
	if (textWidth>totalHeight) textWidth=totalHeight;

	BITMAPINFO dib={sizeof(dib)};
	dib.bmiHeader.biWidth=totalWidth;
	dib.bmiHeader.biHeight=-totalHeight;
	dib.bmiHeader.biPlanes=1;
	dib.bmiHeader.biBitCount=32;
	dib.bmiHeader.biCompression=BI_RGB;

	HDC hdc=CreateCompatibleDC(NULL);
	unsigned int *bits;
	m_Bitmap=CreateDIBSection(hdc,&dib,DIB_RGB_COLORS,(void**)&bits,NULL,0);
	HBITMAP bmp0=(HBITMAP)SelectObject(hdc,m_Bitmap);

	if (opacity==MenuSkin::OPACITY_SOLID)
	{
		RECT rc={0,0,totalWidth,totalHeight};
		SetDCBrushColor(hdc,m_pParent?s_Skin.Submenu_background:s_Skin.Main_background);
		FillRect(hdc,&rc,(HBRUSH)GetStockObject(DC_BRUSH));
	}

	if (bmpSkin)
	{
		// draw the skinned background
		HBITMAP bmp02=(HBITMAP)SelectObject(hdcTemp,bmpSkin);

		RECT rSrc={0,0,slicesX[0]+slicesX[1]+slicesX[2],slicesY[0]+slicesY[1]+slicesY[2]};
		RECT rDst={0,0,textHeight,totalHeight};
		RECT rMargins={slicesX[0],slicesY[0],slicesX[2],slicesY[2]};
		MarginsBlit(hdcTemp,hdc,rSrc,rDst,rMargins,(opacity==MenuSkin::OPACITY_SOLID && b32));

		rSrc.left=rSrc.right;
		rSrc.right+=slicesX[3]+slicesX[4]+slicesX[5];
		rDst.left=rDst.right;
		rDst.right=totalWidth;
		rMargins.left=slicesX[3];
		rMargins.right=slicesX[5];
		MarginsBlit(hdcTemp,hdc,rSrc,rDst,rMargins,(opacity==MenuSkin::OPACITY_SOLID && b32));

		SelectObject(hdcTemp,bmp02);
		SelectObject(hdc,bmp0); // deselect m_Bitmap so all the GDI operations get flushed

		if (s_bRTL)
		{
			// mirror the background image for RTL windows
			for (int y=0;y<totalHeight;y++)
			{
				int yw=y*totalWidth;
				std::reverse(bits+yw,bits+yw+totalWidth);
			}
		}

		// calculate the window region
		m_Region=NULL;
		if (opacity==MenuSkin::OPACITY_REGION || opacity==MenuSkin::OPACITY_GLASS || opacity==MenuSkin::OPACITY_FULLGLASS)
		{
			for (int y=0;y<totalHeight;y++)
			{
				int minx=-1, maxx=-1;
				int yw=y*totalWidth;
				for (int x=0;x<totalWidth;x++)
				{
					if (bits[yw+x]&0xFF000000)
					{
						if (minx==-1) minx=x; // first non-transparent pixel
						if (maxx<x) maxx=x; // last non-transparent pixel
					}
				}
				if (minx>=0)
				{
					maxx++;
					if (s_bRTL && opacity==MenuSkin::OPACITY_REGION)
						minx=totalWidth-minx, maxx=totalWidth-maxx; // in "region" mode mirror the region (again)
					HRGN r=CreateRectRgn(minx,y,maxx,y+1);
					if (!m_Region)
						m_Region=r;
					else
					{
						CombineRgn(m_Region,m_Region,r,RGN_OR);
						DeleteObject(r);
					}
				}
			}
		}

		SelectObject(hdc,m_Bitmap);
	}
	else
	{
		RECT rc={0,0,totalWidth,totalHeight};
		SetDCBrushColor(hdc,m_pParent?s_Skin.Submenu_background:s_Skin.Main_background);
		FillRect(hdc,&rc,(HBRUSH)GetStockObject(DC_BRUSH));
	}

	if (bCaption)
	{
		// draw the title
		BITMAPINFO dib={sizeof(dib)};
		dib.bmiHeader.biWidth=textWidth;
		dib.bmiHeader.biHeight=-textHeight;
		dib.bmiHeader.biPlanes=1;
		dib.bmiHeader.biBitCount=32;
		dib.bmiHeader.biCompression=BI_RGB;

		HDC hdc=CreateCompatibleDC(NULL);
		unsigned int *bits2;
		HBITMAP bmpText=CreateDIBSection(hdcTemp,&dib,DIB_RGB_COLORS,(void**)&bits2,NULL,0);
		HBITMAP bmp02=(HBITMAP)SelectObject(hdcTemp,bmpText);
		{
			RECT rc={0,0,textWidth,textHeight};
			FillRect(hdcTemp,&rc,(HBRUSH)GetStockObject(BLACK_BRUSH));
		}

		RECT rc={s_Skin.Caption_padding.bottom,s_bRTL?s_Skin.Caption_padding.right:s_Skin.Caption_padding.left,textWidth-s_Skin.Caption_padding.top,textHeight-(s_bRTL?s_Skin.Caption_padding.left:s_Skin.Caption_padding.right)};
		if (s_Theme && s_Skin.Caption_glow_size>0)
		{
			// draw the glow
			opts.dwFlags=DTT_COMPOSITED|DTT_TEXTCOLOR|DTT_GLOWSIZE;
			opts.crText=0xFFFFFF;
			opts.iGlowSize=s_Skin.Caption_glow_size;
			DrawThemeTextEx(s_Theme,hdcTemp,0,0,title,-1,DT_VCENTER|DT_NOPREFIX|DT_SINGLELINE,&rc,&opts);
			SelectObject(hdcTemp,bmp02); // deselect bmpText so all the GDI operations get flushed

			// change the glow color
			int gr=(s_Skin.Caption_glow_color)&255;
			int gg=(s_Skin.Caption_glow_color>>8)&255;
			int gb=(s_Skin.Caption_glow_color>>16)&255;
			for (int y=0;y<textHeight;y++)
				for (int x=0;x<textWidth;x++)
				{
					unsigned int &pixel=bits2[y*textWidth+x];
					int a1=(pixel>>24);
					int r1=(pixel>>16)&255;
					int g1=(pixel>>8)&255;
					int b1=(pixel)&255;
					r1=(r1*gr)/255;
					g1=(g1*gg)/255;
					b1=(b1*gb)/255;
					pixel=(a1<<24)|(r1<<16)|(g1<<8)|b1;
				}

				SelectObject(hdcTemp,bmpText);
		}

		// draw the text
		int offset=0;
		if (s_bRTL)
			offset=totalWidth-textHeight;

		if (s_Theme)
		{
			opts.dwFlags=DTT_COMPOSITED|DTT_TEXTCOLOR;
			opts.crText=s_Skin.Caption_text_color;
			DrawThemeTextEx(s_Theme,hdcTemp,0,0,title,-1,DT_VCENTER|DT_NOPREFIX|DT_SINGLELINE,&rc,&opts);
			SelectObject(hdcTemp,bmp02);

			// rotate and copy the text onto the final bitmap. Combine the alpha channels
			for (int y=0;y<textHeight;y++)
				for (int x=0;x<textWidth;x++)
				{
					unsigned int src=bits2[y*textWidth+x];
					int a1=(src>>24);
					int r1=(src>>16)&255;
					int g1=(src>>8)&255;
					int b1=(src)&255;

					unsigned int &dst=bits[(totalHeight-1-x)*totalWidth+y+offset];

					int a2=(dst>>24);
					int r2=(dst>>16)&255;
					int g2=(dst>>8)&255;
					int b2=(dst)&255;

					r2=(r2*(255-a1))/255+r1;
					g2=(g2*(255-a1))/255+g1;
					b2=(b2*(255-a1))/255+b1;
					a2=a1+a2-(a1*a2)/255;

					dst=(a2<<24)|(r2<<16)|(g2<<8)|b2;
				}
		}
		else
		{
			// draw white text on black background
			FillRect(hdcTemp,&rc,(HBRUSH)GetStockObject(BLACK_BRUSH));
			SetTextColor(hdcTemp,0xFFFFFF);
			SetBkMode(hdcTemp,TRANSPARENT);
			DrawText(hdcTemp,title,-1,&rc,DT_VCENTER|DT_NOPREFIX|DT_SINGLELINE);
			SelectObject(hdcTemp,bmp02);

			// rotate and copy the text onto the final bitmap
			// change the text color
			int tr=(s_Skin.Caption_text_color>>16)&255;
			int tg=(s_Skin.Caption_text_color>>8)&255;
			int tb=(s_Skin.Caption_text_color)&255;
			for (int y=0;y<textHeight;y++)
				for (int x=0;x<textWidth;x++)
				{
					unsigned int src=bits2[y*textWidth+x];
					int a1=(src)&255;

					unsigned int &dst=bits[(totalHeight-1-x)*totalWidth+y+offset];

					int a2=(dst>>24);
					int r2=(dst>>16)&255;
					int g2=(dst>>8)&255;
					int b2=(dst)&255;

					r2=(r2*(255-a1)+tr*a1)/255;
					g2=(g2*(255-a1)+tg*a1)/255;
					b2=(b2*(255-a1)+tb*a1)/255;
					a2=a1+a2-(a1*a2)/255;

					dst=(a2<<24)|(r2<<16)|(g2<<8)|b2;
				}
		}

		DeleteObject(bmpText);
	}

	SelectObject(hdcTemp,font0);
	DeleteDC(hdcTemp);

	SelectObject(hdc,bmp0);
	DeleteDC(hdc);

	if (m_pParent)
	{
		m_rContent.left=s_Skin.Submenu_padding.left;
		m_rContent.right=totalWidth-s_Skin.Submenu_padding.right;
		m_rContent.top=s_Skin.Submenu_padding.top;
		m_rContent.bottom=totalHeight-s_Skin.Submenu_padding.bottom;
	}
	else
	{
		m_rContent.left=s_Skin.Main_padding.left+textHeight;
		m_rContent.right=totalWidth-s_Skin.Main_padding.right;
		m_rContent.top=s_Skin.Main_padding.top;
		m_rContent.bottom=totalHeight-s_Skin.Main_padding.bottom;
	}
}

void CMenuContainer::CreateSubmenuRegion( int width, int height )
{
	int totalWidth=s_Skin.Submenu_padding.left+s_Skin.Submenu_padding.right+width;
	int totalHeight=s_Skin.Submenu_padding.top+s_Skin.Submenu_padding.bottom+height;
	m_Region=NULL;
	if (s_Skin.Submenu_opacity!=MenuSkin::OPACITY_REGION && s_Skin.Submenu_opacity!=MenuSkin::OPACITY_GLASS && s_Skin.Submenu_opacity!=MenuSkin::OPACITY_FULLGLASS)
		return;
	if (s_Skin.Submenu_opacity==MenuSkin::OPACITY_SOLID)
		return;
	if (!s_Skin.Submenu_bitmap || !s_Skin.Submenu_bitmap32)
		return;
	BITMAP info;
	GetObject(s_Skin.Submenu_bitmap,sizeof(info),&info);
	const int *slicesX=s_Skin.Submenu_bitmap_slices_X+3;
	const int *slicesY=s_Skin.Submenu_bitmap_slices_Y;
	int slicesX0=slicesX[s_bRTL?2:0];
	int slicesX2=slicesX[s_bRTL?0:2];
	int bmpWidth=slicesX0+slicesX[1]+slicesX2;
	int bmpHeight=slicesY[0]+slicesY[1]+slicesY[2];
	unsigned int *bits=(unsigned int*)info.bmBits;
	{
		for (int y=0;y<totalHeight;y++)
		{
			int yy;
			if (y<=slicesY[0])
				yy=y;
			else if (y>=totalHeight-slicesY[2])
				yy=bmpHeight-totalHeight+y;
			else
				yy=slicesY[0]+((y-slicesY[0])*slicesY[1])/(totalHeight-slicesY[0]-slicesY[2]);
			if (info.bmHeight>0)
				yy=info.bmHeight-yy-1;
			int yw=yy*info.bmWidth;
			int minx=-1, maxx=-1;
			for (int x=0;x<bmpWidth;x++)
			{
				if (bits[yw+x]&0xFF000000)
				{
					if (minx==-1) minx=x; // first non-transparent pixel
					if (maxx<x) maxx=x; // last non-transparent pixel
				}
			}

			if (minx>=0)
			{
				if (minx>=bmpWidth-slicesX2)
					minx+=totalWidth-bmpWidth;
				else if (minx>slicesX0)
					minx=slicesX0+((minx-slicesX0)*(totalWidth-slicesX0-slicesX2))/slicesX[1];

				if (maxx>=bmpWidth-slicesX2)
					maxx+=totalWidth-bmpWidth;
				else if (minx>slicesX0)
					maxx=slicesX0+((maxx-slicesX0)*(totalWidth-slicesX0-slicesX2))/slicesX[1];

				maxx++;
				if (s_bRTL && s_Skin.Submenu_opacity==MenuSkin::OPACITY_REGION)
					minx=totalWidth-minx, maxx=totalWidth-maxx; // in "region" mode mirror the region (again)
				HRGN r=CreateRectRgn(minx,y,maxx,y+1);
				if (!m_Region)
					m_Region=r;
				else
				{
					CombineRgn(m_Region,m_Region,r,RGN_OR);
					DeleteObject(r);
				}
			}
		}
	}
}

void CMenuContainer::DrawBackground( HDC hdc, const RECT &drawRect )
{
	RECT clientRect;
	GetClientRect(&clientRect);

	HDC hdc2=CreateCompatibleDC(hdc);

	// draw the background (bitmap or solid color)
	HBITMAP bmpMain=NULL;
	RECT rMarginsMain, rSrcMain;
	bool bAlphaMain;
	if (m_Bitmap)
	{
		HGDIOBJ bmp0=SelectObject(hdc2,m_Bitmap);
		BitBlt(hdc,0,0,clientRect.right,clientRect.bottom,hdc2,0,0,SRCCOPY);
		SelectObject(hdc2,bmp0);
		bmpMain=m_Bitmap;
		{ RECT rc={clientRect.right,clientRect.bottom,0,0}; rMarginsMain=rc; }
		{ RECT rc={0,0,clientRect.right,clientRect.bottom}; rSrcMain=rc; }
		bAlphaMain=false;
	}
	else if (m_pParent && s_Skin.Submenu_bitmap)
	{
		bAlphaMain=(s_Skin.Submenu_opacity==MenuSkin::OPACITY_SOLID && s_Skin.Submenu_bitmap32);
		if (bAlphaMain)
		{
			SetDCBrushColor(hdc,s_Skin.Submenu_background);
			FillRect(hdc,&drawRect,(HBRUSH)GetStockObject(DC_BRUSH));
		}
		HGDIOBJ bmp0=SelectObject(hdc2,s_Skin.Submenu_bitmap);
		const int *slicesX=s_Skin.Submenu_bitmap_slices_X;
		const int *slicesY=s_Skin.Submenu_bitmap_slices_Y;
		RECT rSrc={0,0,slicesX[3]+slicesX[4]+slicesX[5],slicesY[0]+slicesY[1]+slicesY[2]};
		RECT rMargins={slicesX[3],slicesY[0],slicesX[5],slicesY[2]};
		MarginsBlit(hdc2,hdc,rSrc,clientRect,rMargins,bAlphaMain,s_bRTL);
		SelectObject(hdc2,bmp0);
		bmpMain=s_Skin.Submenu_bitmap;
		rMarginsMain=rMargins;
		rSrcMain=rSrc;
	}
	else
	{
		SetDCBrushColor(hdc,m_pParent?s_Skin.Submenu_background:s_Skin.Main_background);
		FillRect(hdc,&clientRect,(HBRUSH)GetStockObject(DC_BRUSH));
	}

	HBITMAP bmpArrow=NULL;
	bool bArr32=false;
	SIZE arrSize;
	HBITMAP bmpSelection=NULL;
	bool bSel32=false;
	const int *selSlicesX=NULL;
	const int *selSlicesY=NULL;
	COLORREF selColor=0;
	HBITMAP bmpSeparator=NULL;
	bool bSep32=false;
	const int *sepSlicesX=NULL;
	HBITMAP bmpSeparatorV=NULL;
	bool bSepV32=false;
	int sepWidth=0;
	const int *sepSlicesY=NULL;
	HBITMAP bmpPager=NULL;
	bool bPag32=false;
	const int *pagSlicesX=NULL;
	const int *pagSlicesY=NULL;
	HBITMAP bmpPagerArrows=NULL;
	bool bPagArr32=false;
	SIZE pagArrowSize;
	COLORREF *textColors=NULL;
	HIMAGELIST images=(m_Options&CONTAINER_LARGE)?g_IconManager.m_LargeIcons:g_IconManager.m_SmallIcons;
	int iconSize=(m_Options&CONTAINER_LARGE)?g_IconManager.LARGE_ICON_SIZE:g_IconManager.SMALL_ICON_SIZE;
	const RECT &iconPadding=m_pParent?s_Skin.Submenu_icon_padding:s_Skin.Main_icon_padding;
	const RECT &textPadding=m_pParent?s_Skin.Submenu_text_padding:s_Skin.Main_text_padding;
	MenuSkin::TOpacity opacity=m_pParent?s_Skin.Submenu_opacity:s_Skin.Main_opacity;
	int glow=m_pParent?s_Skin.Submenu_glow_size:s_Skin.Main_glow_size;
	if (!s_Theme) glow=0;
	if (m_pParent)
	{
		bmpArrow=s_Skin.Submenu_arrow;
		bArr32=s_Skin.Submenu_arrow32;
		arrSize=s_Skin.Submenu_arrow_Size;
		if (s_Skin.Submenu_selectionColor)
		{
			selColor=s_Skin.Submenu_selection.color;
		}
		else
		{
			bmpSelection=s_Skin.Submenu_selection.bmp;
			bSel32=s_Skin.Submenu_selection32;
			selSlicesX=s_Skin.Submenu_selection_slices_X;
			selSlicesY=s_Skin.Submenu_selection_slices_Y;
		}

		bmpSeparator=s_Skin.Submenu_separator;
		bSep32=s_Skin.Submenu_separator32;
		sepSlicesX=s_Skin.Submenu_separator_slices_X;

		bmpSeparatorV=s_Skin.Submenu_separatorV;
		bSepV32=s_Skin.Submenu_separatorV32;
		sepWidth=s_Skin.Submenu_separatorWidth;
		sepSlicesY=s_Skin.Submenu_separator_slices_Y;

		bmpPager=s_Skin.Submenu_pager;
		bPag32=s_Skin.Submenu_pager32;
		pagSlicesX=s_Skin.Submenu_pager_slices_X;
		pagSlicesY=s_Skin.Submenu_pager_slices_Y;
		bmpPagerArrows=s_Skin.Submenu_pager_arrows;
		bPagArr32=s_Skin.Submenu_pager_arrows32;
		pagArrowSize=s_Skin.Submenu_pager_arrow_Size;

		textColors=s_Skin.Submenu_text_color;
	}
	else
	{
		bmpArrow=s_Skin.Main_arrow;
		bArr32=s_Skin.Main_arrow32;
		arrSize=s_Skin.Main_arrow_Size;
		if (s_Skin.Main_selectionColor)
		{
			selColor=s_Skin.Main_selection.color;
		}
		else
		{
			bmpSelection=s_Skin.Main_selection.bmp;
			bSel32=s_Skin.Main_selection32;
			selSlicesX=s_Skin.Main_selection_slices_X;
			selSlicesY=s_Skin.Main_selection_slices_Y;
		}

		bmpSeparator=s_Skin.Main_separator;
		bSep32=s_Skin.Main_separator32;
		sepSlicesX=s_Skin.Main_separator_slices_X;

		bmpPager=s_Skin.Main_pager;
		bPag32=s_Skin.Main_pager32;
		pagSlicesX=s_Skin.Main_pager_slices_X;
		pagSlicesY=s_Skin.Main_pager_slices_Y;
		bmpPagerArrows=s_Skin.Main_pager_arrows;
		bPagArr32=s_Skin.Main_pager_arrows32;
		pagArrowSize=s_Skin.Main_pager_arrow_Size;

		textColors=s_Skin.Main_text_color;
	}

	HFONT font0=(HFONT)SelectObject(hdc,m_Font);
	SetBkMode(hdc,TRANSPARENT);

	// set clip rectangle for the scrollable items
	int clipTop=m_rContent.top;
	int clipBottom=m_rContent.bottom;
	if (m_bScrollUp)
		clipTop=m_rContent.top+m_ScrollButtonSize;
	if (m_bScrollDown)
		clipBottom=m_rContent.top+m_ScrollHeight-m_ScrollButtonSize;
	if (m_ScrollHeight>0)
		IntersectClipRect(hdc,0,clipTop,clientRect.right,clipBottom);

	// draw items
	for (size_t i=0;;i++)
	{
		if (m_ScrollHeight>0 && (int)i==m_ScrollCount)
		{
			// clean up after the scrollable items
			SelectClipRgn(hdc,NULL);
			if (m_bScrollUp)
			{
				if (glow || opacity==MenuSkin::OPACITY_FULLALPHA || opacity==MenuSkin::OPACITY_FULLGLASS)
				{
					// fix background behind the up button (DrawThemeTextEx may spill onto the tablecloth)
					RECT rc={m_rContent.left,0,m_rContent.right,clipTop};
					if (bAlphaMain || !bmpMain)
					{
						SetDCBrushColor(hdc,m_pParent?s_Skin.Submenu_background:s_Skin.Main_background);
						FillRect(hdc,&rc,(HBRUSH)GetStockObject(DC_BRUSH));
					}
					if (bmpMain)
					{
						HGDIOBJ bmp0=SelectObject(hdc2,bmpMain);
						IntersectClipRect(hdc,m_rContent.left,0,m_rContent.right,clipTop);
						MarginsBlit(hdc2,hdc,rSrcMain,clientRect,rMarginsMain,bAlphaMain,s_bRTL);
						SelectObject(hdc2,bmp0);
						SelectClipRgn(hdc,NULL);
					}
				}

				// draw up button
				RECT rc=m_rContent;
				rc.bottom=clipTop;
				if (bmpPager && bmpPagerArrows)
				{
					// background
					HBITMAP bmp0=(HBITMAP)SelectObject(hdc2,bmpPager);
					RECT rSrc={0,0,pagSlicesX[0]+pagSlicesX[1]+pagSlicesX[2],pagSlicesY[0]+pagSlicesY[1]+pagSlicesY[2]};
					if (m_bScrollUpHot)
						OffsetRect(&rSrc,0,rSrc.bottom);
					RECT rMargins={pagSlicesX[0],pagSlicesY[0],pagSlicesX[2],pagSlicesY[2]};
					MarginsBlit(hdc2,hdc,rSrc,rc,rMargins,bPag32,s_bRTL);

					// arrow
					SelectObject(hdc2,bmpPagerArrows);
					int x=(rc.left+rc.right-pagArrowSize.cx)/2;
					int y=(rc.top+rc.bottom-pagArrowSize.cy)/2;
					if (bPagArr32)
					{
						BLENDFUNCTION func={AC_SRC_OVER,0,255,AC_SRC_ALPHA};
						AlphaBlend(hdc,x,y,pagArrowSize.cx,pagArrowSize.cy,hdc2,m_bScrollUpHot?pagArrowSize.cx:0,0,pagArrowSize.cx,pagArrowSize.cy,func);
					}
					else
						BitBlt(hdc,x,y,pagArrowSize.cx,pagArrowSize.cy,hdc2,m_bScrollUpHot?pagArrowSize.cx:0,0,SRCCOPY);
					SelectObject(hdc2,bmp0);
				}
				else
				{
					if (s_PagerTheme)
						DrawThemeBackground(s_PagerTheme,hdc,SBP_ARROWBTN,(m_bScrollUpHot?ABS_UPHOT:ABS_UPHOVER),&rc,NULL);
					else
						DrawFrameControl(hdc,&rc,DFC_SCROLL,DFCS_SCROLLUP|DFCS_FLAT|(m_bScrollUpHot?DFCS_PUSHED:0));
				}
			}
			if (m_bScrollDown)
			{
				int bottom=clipBottom+m_ItemHeight;
				if (bottom>=clientRect.bottom) bottom=clientRect.bottom;
				if (bottom>clipBottom && (glow || opacity==MenuSkin::OPACITY_FULLALPHA || opacity==MenuSkin::OPACITY_FULLGLASS))
				{
					// fix background behind the up button (DrawThemeTextEx may spill onto the tablecloth)
					RECT rc={m_rContent.left,clipBottom,m_rContent.right,bottom};
					if (bAlphaMain || !bmpMain)
					{
						SetDCBrushColor(hdc,m_pParent?s_Skin.Submenu_background:s_Skin.Main_background);
						FillRect(hdc,&rc,(HBRUSH)GetStockObject(DC_BRUSH));
					}
					if (bmpMain)
					{
						HGDIOBJ bmp0=SelectObject(hdc2,bmpMain);
						IntersectClipRect(hdc,m_rContent.left,clipBottom,m_rContent.right,bottom);
						MarginsBlit(hdc2,hdc,rSrcMain,clientRect,rMarginsMain,bAlphaMain,s_bRTL);
						SelectObject(hdc2,bmp0);
						SelectClipRgn(hdc,NULL);
					}
				}

				// drow down button
				RECT rc=m_rContent;
				rc.bottom=m_rContent.top+m_ScrollHeight;
				rc.top=clipBottom;
				if (bmpPager && bmpPagerArrows)
				{
					// background
					HBITMAP bmp0=(HBITMAP)SelectObject(hdc2,bmpPager);
					RECT rSrc={0,0,pagSlicesX[0]+pagSlicesX[1]+pagSlicesX[2],pagSlicesY[0]+pagSlicesY[1]+pagSlicesY[2]};
					if (m_bScrollDownHot)
						OffsetRect(&rSrc,0,rSrc.bottom);
					RECT rMargins={pagSlicesX[0],pagSlicesY[0],pagSlicesX[2],pagSlicesY[2]};
					MarginsBlit(hdc2,hdc,rSrc,rc,rMargins,bPag32,s_bRTL);

					// arrow
					SelectObject(hdc2,bmpPagerArrows);
					int x=(rc.left+rc.right-pagArrowSize.cx)/2;
					int y=(rc.top+rc.bottom-pagArrowSize.cy)/2;
					if (bPagArr32)
					{
						BLENDFUNCTION func={AC_SRC_OVER,0,255,AC_SRC_ALPHA};
						AlphaBlend(hdc,x,y,pagArrowSize.cx,pagArrowSize.cy,hdc2,m_bScrollDownHot?pagArrowSize.cx:0,pagArrowSize.cy,pagArrowSize.cx,pagArrowSize.cy,func);
					}
					else
						BitBlt(hdc,x,y,pagArrowSize.cx,pagArrowSize.cy,hdc2,m_bScrollDownHot?pagArrowSize.cx:0,pagArrowSize.cy,SRCCOPY);
					SelectObject(hdc2,bmp0);
				}
				else
				{
					if (s_PagerTheme)
						DrawThemeBackground(s_PagerTheme,hdc,SBP_ARROWBTN,(m_bScrollDownHot?ABS_DOWNHOT:ABS_DOWNHOVER),&rc,NULL);
					else
						DrawFrameControl(hdc,&rc,DFC_SCROLL,DFCS_SCROLLDOWN|DFCS_FLAT|(m_bScrollDownHot?DFCS_PUSHED:0));
				}
			}
		}

		if (i>=m_Items.size()) break;
		const MenuItem &item=m_Items[i];
		RECT itemRect=item.itemRect;
		if (m_ScrollHeight>0)
		{
			// ignore offscreen items
			if ((int)i<m_ScrollCount)
			{
				OffsetRect(&itemRect,0,-m_ScrollOffset);
				if (itemRect.bottom<=clipTop) continue;
				if (itemRect.top>=clipBottom) continue;
			}
		}
		{
			RECT q;
			if (!IntersectRect(&q,&drawRect,&itemRect))
				continue;
		}

		if (item.id==MENU_SEPARATOR)
		{
			// draw separator
			if (itemRect.bottom>itemRect.top)
			{
				if (bmpSeparator)
				{
					HBITMAP bmp0=(HBITMAP)SelectObject(hdc2,bmpSeparator);
					RECT rSrc={0,0,sepSlicesX[0]+sepSlicesX[1]+sepSlicesX[2],itemRect.bottom-itemRect.top};
					RECT rMargins={sepSlicesX[0],itemRect.bottom-itemRect.top,sepSlicesX[2],0};
					MarginsBlit(hdc2,hdc,rSrc,itemRect,rMargins,bSep32);
					SelectObject(hdc2,bmp0);
				}
				else
				{
					RECT rc=itemRect;
					if (s_Theme)
					{
						SIZE size;
						if (SUCCEEDED(GetThemePartSize(s_Theme,hdc,TP_SEPARATORVERT,TS_NORMAL,NULL,TS_MIN,&size)))
							OffsetRect(&rc,0,(rc.bottom-rc.top-size.cy)/2);
						DrawThemeBackground(s_Theme,hdc,TP_SEPARATORVERT,TS_NORMAL,&rc,NULL);
					}
					else
					{
						rc.top=rc.bottom=(rc.top+rc.bottom)/2-1;
						DrawEdge(hdc,&rc,EDGE_ETCHED,BF_TOP);
					}
				}
			}
			continue;
		}

		bool bHot=(i==m_HotItem || (m_HotItem==-1 && (i==m_Submenu || i==m_ContextItem)));
		if (bHot)
		{
			// draw selection background
			if (bmpSelection)
			{
				HBITMAP bmp0=(HBITMAP)SelectObject(hdc2,bmpSelection);
				RECT rSrc={0,0,selSlicesX[0]+selSlicesX[1]+selSlicesX[2],selSlicesY[0]+selSlicesY[1]+selSlicesY[2]};
				RECT rMargins={selSlicesX[0],selSlicesY[0],selSlicesX[2],selSlicesY[2]};
				int w=itemRect.right-itemRect.left;
				int h=itemRect.bottom-itemRect.top;
				if (rMargins.left>w) rMargins.left=w;
				if (rMargins.right>w) rMargins.right=w;
				if (rMargins.top>h) rMargins.top=h;
				if (rMargins.bottom>h) rMargins.bottom=h;
				MarginsBlit(hdc2,hdc,rSrc,itemRect,rMargins,bSel32);
				SelectObject(hdc2,bmp0);
			}
			else
			{
				SetDCBrushColor(hdc,selColor);
				FillRect(hdc,&itemRect,(HBRUSH)GetStockObject(DC_BRUSH));
			}
		}

		// draw icon
		ImageList_DrawEx(images,item.icon,hdc,itemRect.left+iconPadding.left,itemRect.top+iconPadding.top+m_IconTopOffset,0,0,CLR_NONE,CLR_NONE,ILD_NORMAL);

		// draw text
		COLORREF color;
		if (item.id==MENU_EMPTY)
			color=textColors[bHot?3:2];
		else
			color=textColors[bHot?1:0];
		RECT rc={itemRect.left+iconSize+iconPadding.left+iconPadding.right+textPadding.left,itemRect.top+m_TextTopOffset,itemRect.right-ARROW_SIZE-textPadding.right,itemRect.bottom-m_TextTopOffset};
		DWORD flags=DT_VCENTER|DT_SINGLELINE|DT_END_ELLIPSIS;
		if (item.id==MENU_NO)
			flags|=DT_NOPREFIX;
		else if (!s_bKeyboardCues)
			flags|=DT_HIDEPREFIX;
		if (s_Theme)
		{
			DTTOPTS opts={sizeof(opts),DTT_TEXTCOLOR};
			opts.crText=color;
			if (glow || opacity==MenuSkin::OPACITY_FULLALPHA || opacity==MenuSkin::OPACITY_FULLGLASS)
				opts.dwFlags|=DTT_COMPOSITED;
			if (glow)
			{
				opts.dwFlags|=DTT_GLOWSIZE;
				opts.iGlowSize=glow;
			}
			DrawThemeTextEx(s_Theme,hdc,0,0,item.name,item.name.GetLength(),flags,&rc,&opts);
		}
		else
		{
			SetTextColor(hdc,color);
			DrawText(hdc,item.name,item.name.GetLength(),&rc,flags);
		}

		if (item.bFolder)
		{
			// draw the sub-menu arrows
			if (bmpArrow)
			{
				int x=itemRect.right-4-arrSize.cx;
				int y=(itemRect.top+itemRect.bottom-arrSize.cy)/2;
				HBITMAP bmp0=(HBITMAP)SelectObject(hdc2,bmpArrow);
				if (bArr32)
				{
					BLENDFUNCTION func={AC_SRC_OVER,0,255,AC_SRC_ALPHA};
					AlphaBlend(hdc,x,y,arrSize.cx,arrSize.cy,hdc2,0,bHot?arrSize.cy:0,arrSize.cx,arrSize.cy,func);
				}
				else
				{
					BitBlt(hdc,x,y,arrSize.cx,arrSize.cy,hdc2,0,bHot?arrSize.cy:0,SRCCOPY);
				}
				SelectObject(hdc2,bmp0);
			}
			else
			{
				int x=itemRect.right-8;
				int y=(itemRect.top+itemRect.bottom-6)/2;
				HBITMAP bmp0=(HBITMAP)SelectObject(hdc2,bHot?m_ArrowsBitmapSel:m_ArrowsBitmap);
				BLENDFUNCTION func={AC_SRC_OVER,0,255,AC_SRC_ALPHA};
				AlphaBlend(hdc,x,y,4,7,hdc2,s_bRTL?0:10,0,4,7,func);
				SelectObject(hdc2,bmp0);
			}
		}
	}

	// draw vertical separators
	if (bmpSeparatorV && m_ColumnOffsets.size()>1)
	{
		HBITMAP bmp0=(HBITMAP)SelectObject(hdc2,bmpSeparatorV);
		RECT rSrc={0,0,sepWidth,sepSlicesY[0]+sepSlicesY[1]+sepSlicesY[2]};
		RECT rMargins={0,sepSlicesY[0],0,sepSlicesY[2]};
		for (size_t i=1;i<m_ColumnOffsets.size();i++)
		{
			int x=m_rContent.left+m_ColumnOffsets[i];
			RECT rc={x-sepWidth,m_rContent.top,x,m_rContent.bottom};
			MarginsBlit(hdc2,hdc,rSrc,rc,rMargins,bSepV32);
		}
		SelectObject(hdc2,bmp0);
	}

	// draw insert mark
	{
		RECT rc;
		if (GetInsertRect(rc))
		{
			HBITMAP bmp0=(HBITMAP)SelectObject(hdc2,m_ArrowsBitmap);
			RECT rSrc={s_bRTL?9:0,0,s_bRTL?18:9,6};
			RECT rMargins={3,0,3,0};
			MarginsBlit(hdc2,hdc,rSrc,rc,rMargins,true);
			SelectObject(hdc2,bmp0);
		}
	}

	SelectObject(hdc,font0);

	DeleteDC(hdc2);
}

LRESULT CMenuContainer::OnPaint( UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled )
{
	// handles both WM_PAINT and WM_PRINTCLIENT
	PAINTSTRUCT ps;
	HDC hdc;
	if (uMsg==WM_PRINTCLIENT)
	{
		hdc=(HDC)wParam;
		GetClientRect(&ps.rcPaint);
	}
	else
	{
		hdc=BeginPaint(&ps);
	}
	MenuSkin::TOpacity opacity=(m_pParent?s_Skin.Submenu_opacity:s_Skin.Main_opacity);
	if ((opacity==MenuSkin::OPACITY_GLASS || opacity==MenuSkin::OPACITY_FULLGLASS) && ((!m_pParent && m_Bitmap) || (m_pParent && s_Skin.Submenu_bitmap)))
	{
		DWM_BLURBEHIND blur={DWM_BB_ENABLE|DWM_BB_BLURREGION,TRUE,m_Region,FALSE};
		DwmEnableBlurBehindWindow(m_hWnd,&blur);
	}
	BP_PAINTPARAMS paintParams={sizeof(paintParams)};
	paintParams.dwFlags=BPPF_ERASE;

	HDC hdcPaint=NULL;
	HPAINTBUFFER hBufferedPaint=BeginBufferedPaint(hdc,&ps.rcPaint,BPBF_TOPDOWNDIB,&paintParams,&hdcPaint);
	if (hdcPaint)
	{
		DrawBackground(hdcPaint,ps.rcPaint);
		if (opacity==MenuSkin::OPACITY_GLASS || opacity==MenuSkin::OPACITY_ALPHA)
		{
			RECT rc;
			IntersectRect(&rc,&ps.rcPaint,&m_rContent);
			BufferedPaintSetAlpha(hBufferedPaint,&rc,255);
		}
		EndBufferedPaint(hBufferedPaint,TRUE);
	}

	if (uMsg!=WM_PRINTCLIENT)
		EndPaint(&ps);

	return 0;
}
