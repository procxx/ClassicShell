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
#define SECURITY_WIN32
#include <Security.h>
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
void CMenuContainer::CreateBackground( int width1, int width2, int height1, int height2 )
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

	HBITMAP bmpSkin=m_bSubMenu?s_Skin.Submenu_bitmap:s_Skin.Main_bitmap;
	bool b32=m_bSubMenu?s_Skin.Submenu_bitmap32:s_Skin.Main_bitmap32;
	const int *slicesX=m_bSubMenu?s_Skin.Submenu_bitmap_slices_X:s_Skin.Main_bitmap_slices_X;
	const int *slicesY=m_bSubMenu?s_Skin.Submenu_bitmap_slices_Y:s_Skin.Main_bitmap_slices_Y;
	bool bCaption=(slicesX[1]>0);
	MenuSkin::TOpacity opacity=m_bSubMenu?s_Skin.Submenu_opacity:s_Skin.Main_opacity;

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
	if (!m_bSubMenu)
	{
		textWidth=rc.right+s_Skin.Caption_padding.top+s_Skin.Caption_padding.bottom;
		textHeight=rc.bottom+s_Skin.Caption_padding.left+s_Skin.Caption_padding.right;
	}
	int total=slicesX[0]+slicesX[2];
	if (textHeight<total) textHeight=total;

	int totalWidth1=textHeight+width1;
	totalWidth1+=m_bSubMenu?(s_Skin.Submenu_padding.left+s_Skin.Submenu_padding.right):(s_Skin.Main_padding.left+s_Skin.Main_padding.right);
	if (totalWidth1<total) totalWidth1=total;
	int totalWidth2=0;
	if (m_bTwoColumns)
		totalWidth2=width2+s_Skin.Main_padding2.left+s_Skin.Main_padding2.right;
	int totalWidth=totalWidth1+totalWidth2;

	total+=textHeight+slicesX[3]+slicesX[5];

	int totalHeight=height1+(m_bSubMenu?(s_Skin.Submenu_padding.top+s_Skin.Submenu_padding.bottom):(s_Skin.Main_padding.top+s_Skin.Main_padding.bottom));
	if (m_bTwoColumns)
	{
		int totalHeight2=height2+s_Skin.Main_padding2.top+s_Skin.Main_padding2.bottom;
		if (totalHeight<totalHeight2) totalHeight=totalHeight2;
	}

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
		SetDCBrushColor(hdc,m_bSubMenu?s_Skin.Submenu_background:s_Skin.Main_background);
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
		rDst.right=totalWidth1;
		rMargins.left=slicesX[3];
		rMargins.right=slicesX[5];
		MarginsBlit(hdcTemp,hdc,rSrc,rDst,rMargins,(opacity==MenuSkin::OPACITY_SOLID && b32));

		if (totalWidth2>0)
		{
			rSrc.left=rSrc.right;
			rSrc.right+=slicesX[6]+slicesX[7]+slicesX[8];
			rDst.left=rDst.right;
			rDst.right+=totalWidth2;
			rMargins.left=slicesX[6];
			rMargins.right=slicesX[8];
			MarginsBlit(hdcTemp,hdc,rSrc,rDst,rMargins,(opacity==MenuSkin::OPACITY_SOLID && b32));
		}

		if (width2 && s_Skin.Main_separatorV)
		{
			SelectObject(hdcTemp,s_Skin.Main_separatorV);
			RECT rSrc2={0,0,s_Skin.Main_separatorWidth,s_Skin.Main_separator_slices_Y[0]+s_Skin.Main_separator_slices_Y[1]+s_Skin.Main_separator_slices_Y[2]};
			RECT rDst2={totalWidth1,s_Skin.Main_padding.top,totalWidth1+s_Skin.Main_separatorWidth,totalHeight-s_Skin.Main_padding.bottom};
			RECT rMargins2={0,s_Skin.Main_separator_slices_Y[0],0,s_Skin.Main_separator_slices_Y[2]};
			MarginsBlit(hdcTemp,hdc,rSrc2,rDst2,rMargins2,s_Skin.Main_separatorV32);
		}

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

		SelectObject(hdc,m_Bitmap);
	}
	else
	{
		RECT rc={0,0,totalWidth,totalHeight};
		SetDCBrushColor(hdc,m_bSubMenu?s_Skin.Submenu_background:s_Skin.Main_background);
		FillRect(hdc,&rc,(HBRUSH)GetStockObject(DC_BRUSH));
		if (width2)
		{
			if (s_Skin.Main_separatorV)
			{
				HBITMAP bmp02=(HBITMAP)SelectObject(hdcTemp,s_Skin.Main_separatorV);
				RECT rSrc2={0,0,s_Skin.Main_separatorWidth,s_Skin.Main_separator_slices_Y[0]+s_Skin.Main_separator_slices_Y[1]+s_Skin.Main_separator_slices_Y[2]};
				RECT rDst2={totalWidth1,s_Skin.Main_padding.top,totalWidth1+s_Skin.Main_separatorWidth,totalHeight-s_Skin.Main_padding.bottom};
				RECT rMargins2={0,s_Skin.Main_separator_slices_Y[0],0,s_Skin.Main_separator_slices_Y[2]};
				MarginsBlit(hdcTemp,hdc,rSrc2,rDst2,rMargins2,s_Skin.Main_separatorV32);
				SelectObject(hdcTemp,bmp02);
			}
			else
			{
				rc.left=rc.right=totalWidth1;
				rc.top+=s_Skin.Main_padding.top;
				rc.bottom-=s_Skin.Main_padding.bottom;
				DrawEdge(hdc,&rc,EDGE_ETCHED,BF_LEFT);
			}
		}
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
		SelectObject(hdcTemp,font0);
	}

	if (s_Skin.User_image_size)
	{
		// draw user image
		HBITMAP userPicture=NULL;
		HMODULE hShell32=GetModuleHandle(L"shell32.dll");
		typedef HRESULT (__stdcall*tSHGetUserPicturePath)(LPCWSTR, UINT, LPWSTR, ULONG);
		tSHGetUserPicturePath SHGetUserPicturePath=(tSHGetUserPicturePath)GetProcAddress(hShell32,MAKEINTRESOURCEA(261));
		if (SHGetUserPicturePath)
		{
			wchar_t path[_MAX_PATH];
			path[0]=0;
			SHGetUserPicturePath(NULL,0x80000000,path,_countof(path));
			userPicture=(HBITMAP)LoadImage(NULL,path,IMAGE_BITMAP,0,0,LR_LOADFROMFILE);
			if (userPicture)
			{
				// resize the user bitmap to the required size using HALFTONE stretch mode. LoadImage uses COLORONCOLOR internally, which is not very good for the user image
				BITMAP info;
				GetObject(userPicture,sizeof(info),&info);

				HDC hdc3=CreateCompatibleDC(hdcTemp);
				BITMAPINFO dib2={sizeof(dib2)};
				dib2.bmiHeader.biWidth=s_Skin.User_image_size;
				dib2.bmiHeader.biHeight=s_Skin.User_image_size;
				dib2.bmiHeader.biPlanes=1;
				dib2.bmiHeader.biBitCount=24;
				dib2.bmiHeader.biCompression=BI_RGB;
				HBITMAP userPicture2=CreateDIBSection(hdc3,&dib2,DIB_RGB_COLORS,NULL,NULL,0);
				HBITMAP bmp03=(HBITMAP)SelectObject(hdc3,userPicture2);

				HBITMAP bmp02=(HBITMAP)SelectObject(hdcTemp,userPicture);
				SetStretchBltMode(hdc3,HALFTONE);
				StretchBlt(hdc3,0,0,s_Skin.User_image_size,s_Skin.User_image_size,hdcTemp,0,0,info.bmWidth,info.bmHeight,SRCCOPY);

				SelectObject(hdc3,bmp03);
				SelectObject(hdcTemp,bmp02);
				DeleteDC(hdc3);
				DeleteObject(userPicture);
				userPicture=userPicture2;
			}
		}
		if (userPicture)
		{
			// draw user picture
			SIZE frameSize;
			if (s_Skin.User_bitmap)
			{
				BITMAP info;
				GetObject(s_Skin.User_bitmap,sizeof(info),&info);
				frameSize.cx=info.bmWidth;
				frameSize.cy=info.bmHeight;
			}
			else
			{
				frameSize.cx=s_Skin.User_image_size+s_Skin.User_image_offset.x*2;
				frameSize.cy=s_Skin.User_image_size+s_Skin.User_image_offset.y*2;
			}
			POINT pos=s_Skin.User_frame_position;
			if (pos.x==MenuSkin::USER_CENTER)
				pos.x=(totalWidth-frameSize.cx)/2;
			else if (pos.x==MenuSkin::USER_CENTER1)
				pos.x=(totalWidth1+textHeight-frameSize.cx)/2;
			else if (pos.x==MenuSkin::USER_CENTER2)
			{
				if (totalWidth2>0)
					pos.x=totalWidth1+(totalWidth2-frameSize.cx)/2;
				else
					pos.x=(totalWidth-frameSize.cx)/2;
			}

			if (pos.x<0) pos.x+=totalWidth-frameSize.cx;
			if (pos.y<0) pos.y+=totalHeight-frameSize.cy;

			if (s_bRTL)
				pos.x=totalWidth-frameSize.cx-pos.x;
			pos.x+=s_Skin.User_image_offset.x;
			pos.y+=s_Skin.User_image_offset.y;
			HBITMAP bmp02=(HBITMAP)SelectObject(hdcTemp,userPicture);
			unsigned int alpha=s_Skin.User_image_alpha;
			if (alpha==255)
			{
				BitBlt(hdc,pos.x,pos.y,s_Skin.User_image_size,s_Skin.User_image_size,hdcTemp,0,0,SRCCOPY);
			}
			else
			{
				BLENDFUNCTION func={AC_SRC_OVER,0,alpha,0};
				AlphaBlend(hdc,pos.x,pos.y,s_Skin.User_image_size,s_Skin.User_image_size,hdcTemp,0,0,s_Skin.User_image_size,s_Skin.User_image_size,func);
			}

			if (s_bRTL)
				m_rUser1.left=totalWidth-pos.x-s_Skin.User_image_size;
			else
				m_rUser1.left=pos.x;
			m_rUser1.right=m_rUser1.left+s_Skin.User_image_size;
			m_rUser1.top=pos.y;
			m_rUser1.bottom=pos.y+s_Skin.User_image_size;

			if (opacity!=MenuSkin::OPACITY_SOLID)
			{
				// set to opaque
				SelectObject(hdc,bmp0); // deselect m_Bitmap so all the GDI operations get flushed
				unsigned int *bits2=bits+pos.y*totalWidth+pos.x;
				alpha<<=24;
				for (int y=0;y<s_Skin.User_image_size;y++,bits2+=totalWidth)
					for (int x=0;x<s_Skin.User_image_size;x++)
						bits2[x]=alpha|(bits2[x]&0xFFFFFF);
				SelectObject(hdc,m_Bitmap);
			}

			// draw frame
			pos.x-=s_Skin.User_image_offset.x;
			pos.y-=s_Skin.User_image_offset.y;
			if (s_Skin.User_bitmap)
			{
				SelectObject(hdcTemp,s_Skin.User_bitmap);
				BLENDFUNCTION func={AC_SRC_OVER,0,255,AC_SRC_ALPHA};
				AlphaBlend(hdc,pos.x,pos.y,frameSize.cx,frameSize.cy,hdcTemp,0,0,frameSize.cx,frameSize.cy,func);
			}
			else
			{
				RECT rc={pos.x,pos.y,pos.x+frameSize.cx,pos.y+frameSize.cy};
				DrawEdge(hdc,&rc,EDGE_BUMP,BF_RECT);
			}
			SelectObject(hdcTemp,bmp02);
			DeleteObject(userPicture);
		}
	}

	if (s_Skin.User_name_position.left!=s_Skin.User_name_position.right)
	{
		RECT rc0;
		int x0=0, x1=totalWidth;
		if (s_Skin.User_name_align==MenuSkin::NAME_CENTER1 || s_Skin.User_name_align==MenuSkin::NAME_LEFT1 || s_Skin.User_name_align==MenuSkin::NAME_RIGHT1)
			x1=totalWidth1;
		else if (s_Skin.User_name_align==MenuSkin::NAME_CENTER2 || s_Skin.User_name_align==MenuSkin::NAME_LEFT2 || s_Skin.User_name_align==MenuSkin::NAME_RIGHT2)
			x0=totalWidth1;

		if (s_Skin.User_name_position.left<0)
			rc0.left=x1+s_Skin.User_name_position.left;
		else
			rc0.left=x0+s_Skin.User_name_position.left;

		if (s_Skin.User_name_position.right<0)
			rc0.right=x1+s_Skin.User_name_position.right;
		else
			rc0.right=x0+s_Skin.User_name_position.right;

		rc0.top=s_Skin.User_name_position.top;
		if (rc0.top<0) rc0.top+=totalHeight;
		rc0.bottom=s_Skin.User_name_position.bottom;
		if (rc0.bottom<0) rc0.bottom+=totalHeight;

		m_rUser2=rc0;

		wchar_t name[256];
		const wchar_t *str=FindSetting("MenuUsername");
		if (str)
		{
			Strcpy(name,_countof(name),str);
			DoEnvironmentSubst(name,_countof(name));
		}
		else
		{
			ULONG size=_countof(name);
			if (!GetUserNameEx(NameDisplay,name,&size))
			{
				// GetUserNameEx may fail (for example on Home editions). use the login name
				size=_countof(name);
				GetUserName(name,&size);
			}
		}

		if (Strlen(name)>0)
		{
			int nameWidth=rc0.right-rc0.left;
			int nameHeight=rc0.bottom-rc0.top;
			RECT rc={0,0,nameWidth,nameHeight};

			// draw the title
			BITMAPINFO dib={sizeof(dib)};
			dib.bmiHeader.biWidth=nameWidth;
			dib.bmiHeader.biHeight=-nameHeight;
			dib.bmiHeader.biPlanes=1;
			dib.bmiHeader.biBitCount=32;
			dib.bmiHeader.biCompression=BI_RGB;

			font0=(HFONT)SelectObject(hdcTemp,s_Skin.User_font);

			unsigned int *bits2;
			HBITMAP bmpText=CreateDIBSection(hdcTemp,&dib,DIB_RGB_COLORS,(void**)&bits2,NULL,0);
			HBITMAP bmp02=(HBITMAP)SelectObject(hdcTemp,bmpText);
			FillRect(hdcTemp,&rc,(HBRUSH)GetStockObject(BLACK_BRUSH));

			DWORD align=DT_CENTER;
			if (s_Skin.User_name_align==MenuSkin::NAME_LEFT || s_Skin.User_name_align==MenuSkin::NAME_LEFT1 || s_Skin.User_name_align==MenuSkin::NAME_LEFT2)
				align=s_bRTL?DT_RIGHT:DT_LEFT;
			else if (s_Skin.User_name_align==MenuSkin::NAME_RIGHT || s_Skin.User_name_align==MenuSkin::NAME_RIGHT1 || s_Skin.User_name_align==MenuSkin::NAME_RIGHT2)
				align=s_bRTL?DT_LEFT:DT_RIGHT;
			if (s_Theme && s_Skin.User_glow_size>0)
			{
				InflateRect(&rc,-s_Skin.User_glow_size,-s_Skin.User_glow_size);
				// draw the glow
				opts.dwFlags=DTT_COMPOSITED|DTT_TEXTCOLOR|DTT_GLOWSIZE;
				opts.crText=0xFFFFFF;
				opts.iGlowSize=s_Skin.User_glow_size;
				DrawThemeTextEx(s_Theme,hdcTemp,0,0,name,-1,align|DT_VCENTER|DT_NOPREFIX|DT_SINGLELINE|DT_END_ELLIPSIS|DT_NOCLIP,&rc,&opts);
				SelectObject(hdcTemp,bmp02); // deselect bmpText so all the GDI operations get flushed

				// change the glow color
				int gr=(s_Skin.User_glow_color)&255;
				int gg=(s_Skin.User_glow_color>>8)&255;
				int gb=(s_Skin.User_glow_color>>16)&255;
				for (int y=0;y<nameHeight;y++)
					for (int x=0;x<nameWidth;x++)
					{
						unsigned int &pixel=bits2[y*nameWidth+x];
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
			int offset=rc0.top*totalWidth+rc0.left;
			if (s_bRTL)
				offset=rc0.top*totalWidth+totalWidth-rc0.right;

			if (s_Theme)
			{
				opts.dwFlags=DTT_COMPOSITED|DTT_TEXTCOLOR;
				opts.crText=s_Skin.User_text_color;
				DrawThemeTextEx(s_Theme,hdcTemp,0,0,name,-1,align|DT_VCENTER|DT_NOPREFIX|DT_SINGLELINE|DT_END_ELLIPSIS|DT_NOCLIP,&rc,&opts);
				SelectObject(hdcTemp,bmp02);

				// copy the text onto the final bitmap. Combine the alpha channels
				for (int y=0;y<nameHeight;y++)
					for (int x=0;x<nameWidth;x++)
					{
						unsigned int src=bits2[y*nameWidth+x];
						int a1=(src>>24);
						int r1=(src>>16)&255;
						int g1=(src>>8)&255;
						int b1=(src)&255;

						unsigned int &dst=bits[y*totalWidth+x+offset];

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
				SetTextColor(hdcTemp,0xFFFFFF);
				SetBkMode(hdcTemp,TRANSPARENT);
				DrawText(hdcTemp,name,-1,&rc,align|DT_VCENTER|DT_NOPREFIX|DT_SINGLELINE|DT_END_ELLIPSIS|DT_NOCLIP);
				SelectObject(hdcTemp,bmp02);

				// copy the text onto the final bitmap
				// change the text color
				int tr=(s_Skin.User_text_color>>16)&255;
				int tg=(s_Skin.User_text_color>>8)&255;
				int tb=(s_Skin.User_text_color)&255;
				for (int y=0;y<nameHeight;y++)
					for (int x=0;x<nameWidth;x++)
					{
						unsigned int src=bits2[y*nameWidth+x];
						int a1=(src)&255;

						unsigned int &dst=bits[y*totalWidth+x+offset];

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
			SelectObject(hdcTemp,font0);
		}
	}

	DeleteDC(hdcTemp);

	SelectObject(hdc,bmp0);
	DeleteDC(hdc);

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

	if (m_bSubMenu)
	{
		m_rContent.left=s_Skin.Submenu_padding.left;
		m_rContent.right=totalWidth1-s_Skin.Submenu_padding.right;
		m_rContent.top=s_Skin.Submenu_padding.top;
		m_rContent.bottom=totalHeight-s_Skin.Submenu_padding.bottom;
	}
	else
	{
		m_rContent.left=s_Skin.Main_padding.left+textHeight;
		m_rContent.right=totalWidth1-s_Skin.Main_padding.right;
		m_rContent.top=s_Skin.Main_padding.top;
		m_rContent.bottom=m_rContent.top+height1;
		if (totalWidth2>0)
		{
			m_rContent2.left=m_rContent.right+s_Skin.Main_padding.right+s_Skin.Main_padding2.left;
			m_rContent2.right=totalWidth-s_Skin.Main_padding2.right;
			m_rContent2.top=s_Skin.Main_padding2.top;
			m_rContent2.bottom=m_rContent2.top+height2;
		}
	}
}

void CMenuContainer::CreateSubmenuRegion( int width, int height )
{
	int totalWidth=s_Skin.Submenu_padding.left+s_Skin.Submenu_padding.right+width;
	int totalHeight=s_Skin.Submenu_padding.top+s_Skin.Submenu_padding.bottom+height;
	m_Region=NULL;
	if (s_Skin.Submenu_opacity!=MenuSkin::OPACITY_REGION && s_Skin.Submenu_opacity!=MenuSkin::OPACITY_GLASS && s_Skin.Submenu_opacity!=MenuSkin::OPACITY_FULLGLASS)
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
	else if (m_bSubMenu && s_Skin.Submenu_bitmap)
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
		SetDCBrushColor(hdc,m_bSubMenu?s_Skin.Submenu_background:s_Skin.Main_background);
		FillRect(hdc,&clientRect,(HBRUSH)GetStockObject(DC_BRUSH));
	}

	COLORREF *textColors[2]={NULL,NULL};
	HBITMAP bmpArrow[2]={NULL,NULL};
	bool bArr32[2]={false,false};
	SIZE arrSize[2];
	HBITMAP bmpSelection[2]={NULL,NULL};
	bool bSel32[2]={false,false};
	const int *selSlicesX[2]={NULL,NULL};
	const int *selSlicesY[2]={NULL,NULL};
	COLORREF selColor[2]={0,0};
	HBITMAP bmpSeparator[2]={NULL,NULL};
	bool bSep32[2]={false,false};
	const int *sepSlicesX[2]={NULL,NULL};
	const RECT iconPadding[2]={m_bSubMenu?s_Skin.Submenu_icon_padding:s_Skin.Main_icon_padding,s_Skin.Main_icon_padding2};
	const RECT textPadding[2]={m_bSubMenu?s_Skin.Submenu_text_padding:s_Skin.Main_text_padding,s_Skin.Main_text_padding2};
	const SIZE arrPadding[2]={m_bSubMenu?s_Skin.Submenu_arrow_padding:s_Skin.Main_arrow_padding,s_Skin.Main_arrow_padding2};
	MenuSkin::TOpacity opacity[2]={m_bSubMenu?s_Skin.Submenu_opacity:s_Skin.Main_opacity,s_Skin.Main_opacity2};
	int glow[2]={m_bSubMenu?s_Skin.Submenu_glow_size:s_Skin.Main_glow_size,s_Skin.Main_glow_size2};

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
	HIMAGELIST images=(m_Options&CONTAINER_LARGE)?g_IconManager.m_LargeIcons:g_IconManager.m_SmallIcons;
	int iconSize=(m_Options&CONTAINER_LARGE)?g_IconManager.LARGE_ICON_SIZE:g_IconManager.SMALL_ICON_SIZE;

	if (!s_Theme) glow[0]=glow[1]=0;

	if (m_bSubMenu)
	{
		textColors[0]=s_Skin.Submenu_text_color;
		bmpArrow[0]=s_Skin.Submenu_arrow;
		bArr32[0]=s_Skin.Submenu_arrow32;
		arrSize[0]=s_Skin.Submenu_arrow_Size;
		if (s_Skin.Submenu_selectionColor)
		{
			selColor[0]=s_Skin.Submenu_selection.color;
		}
		else
		{
			bmpSelection[0]=s_Skin.Submenu_selection.bmp;
			bSel32[0]=s_Skin.Submenu_selection32;
			selSlicesX[0]=s_Skin.Submenu_selection_slices_X;
			selSlicesY[0]=s_Skin.Submenu_selection_slices_Y;
		}

		bmpSeparator[0]=s_Skin.Submenu_separator;
		bSep32[0]=s_Skin.Submenu_separator32;
		sepSlicesX[0]=s_Skin.Submenu_separator_slices_X;

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
	}
	else
	{
		textColors[0]=s_Skin.Main_text_color;
		bmpArrow[0]=s_Skin.Main_arrow;
		bArr32[0]=s_Skin.Main_arrow32;
		arrSize[0]=s_Skin.Main_arrow_Size;
		if (s_Skin.Main_selectionColor)
		{
			selColor[0]=s_Skin.Main_selection.color;
		}
		else
		{
			bmpSelection[0]=s_Skin.Main_selection.bmp;
			bSel32[0]=s_Skin.Main_selection32;
			selSlicesX[0]=s_Skin.Main_selection_slices_X;
			selSlicesY[0]=s_Skin.Main_selection_slices_Y;
		}
		bmpSeparator[0]=s_Skin.Main_separator;
		bSep32[0]=s_Skin.Main_separator32;
		sepSlicesX[0]=s_Skin.Main_separator_slices_X;

		if (m_bTwoColumns)
		{
			textColors[1]=s_Skin.Main_text_color2;
			bmpArrow[1]=s_Skin.Main_arrow2;
			bArr32[1]=s_Skin.Main_arrow232;
			arrSize[1]=s_Skin.Main_arrow_Size2;
			if (s_Skin.Main_selectionColor2)
			{
				selColor[1]=s_Skin.Main_selection2.color;
			}
			else
			{
				bmpSelection[1]=s_Skin.Main_selection2.bmp;
				bSel32[1]=s_Skin.Main_selection232;
				selSlicesX[1]=s_Skin.Main_selection_slices_X2;
				selSlicesY[1]=s_Skin.Main_selection_slices_Y2;
			}

			bmpSeparator[1]=s_Skin.Main_separator2;
			bSep32[1]=s_Skin.Main_separator232;
			sepSlicesX[1]=s_Skin.Main_separator_slices_X2;
		}

		bmpPager=s_Skin.Main_pager;
		bPag32=s_Skin.Main_pager32;
		pagSlicesX=s_Skin.Main_pager_slices_X;
		pagSlicesY=s_Skin.Main_pager_slices_Y;
		bmpPagerArrows=s_Skin.Main_pager_arrows;
		bPagArr32=s_Skin.Main_pager_arrows32;
		pagArrowSize=s_Skin.Main_pager_arrow_Size;
	}

	HFONT font0=(HFONT)SelectObject(hdc,m_Font[0]);
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
				if (glow[0] || opacity[0]==MenuSkin::OPACITY_FULLALPHA || opacity[0]==MenuSkin::OPACITY_FULLGLASS)
				{
					// fix background behind the up button (DrawThemeTextEx may spill onto the tablecloth)
					RECT rc={m_rContent.left,0,m_rContent.right,clipTop};
					if (bAlphaMain || !bmpMain)
					{
						SetDCBrushColor(hdc,m_bSubMenu?s_Skin.Submenu_background:s_Skin.Main_background);
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
				int bottom=clipBottom+m_ItemHeight[0];
				if (bottom>=clientRect.bottom) bottom=clientRect.bottom;
				if (bottom>clipBottom && (glow[0] || opacity[0]==MenuSkin::OPACITY_FULLALPHA || opacity[0]==MenuSkin::OPACITY_FULLGLASS))
				{
					// fix background behind the up button (DrawThemeTextEx may spill onto the tablecloth)
					RECT rc={m_rContent.left,clipBottom,m_rContent.right,bottom};
					if (bAlphaMain || !bmpMain)
					{
						SetDCBrushColor(hdc,m_bSubMenu?s_Skin.Submenu_background:s_Skin.Main_background);
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

				// draw down button
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

		int index=(m_bTwoColumns && item.column==1)?1:0;
		if (index==1 && (i==0 || m_Items[i-1].column==0))
		{
			SelectObject(hdc,m_Font[1]);
		}

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
				if (bmpSeparator[index])
				{
					HBITMAP bmp0=(HBITMAP)SelectObject(hdc2,bmpSeparator[index]);
					RECT rSrc={0,0,sepSlicesX[index][0]+sepSlicesX[index][1]+sepSlicesX[index][2],itemRect.bottom-itemRect.top};
					RECT rMargins={sepSlicesX[index][0],itemRect.bottom-itemRect.top,sepSlicesX[index][2],0};
					MarginsBlit(hdc2,hdc,rSrc,itemRect,rMargins,bSep32[index]);
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
			if (bmpSelection[index])
			{
				HBITMAP bmp0=(HBITMAP)SelectObject(hdc2,bmpSelection[index]);
				RECT rSrc={0,0,selSlicesX[index][0]+selSlicesX[index][1]+selSlicesX[index][2],selSlicesY[index][0]+selSlicesY[index][1]+selSlicesY[index][2]};
				RECT rMargins={selSlicesX[index][0],selSlicesY[index][0],selSlicesX[index][2],selSlicesY[index][2]};
				int w=itemRect.right-itemRect.left;
				int h=itemRect.bottom-itemRect.top;
				if (rMargins.left>w) rMargins.left=w;
				if (rMargins.right>w) rMargins.right=w;
				if (rMargins.top>h) rMargins.top=h;
				if (rMargins.bottom>h) rMargins.bottom=h;
				MarginsBlit(hdc2,hdc,rSrc,itemRect,rMargins,bSel32[index]);
				SelectObject(hdc2,bmp0);
			}
			else
			{
				SetDCBrushColor(hdc,selColor[index]);
				FillRect(hdc,&itemRect,(HBRUSH)GetStockObject(DC_BRUSH));
			}
		}

		// draw icon
		ImageList_DrawEx(images,item.icon,hdc,itemRect.left+iconPadding[index].left,itemRect.top+iconPadding[index].top+m_IconTopOffset[index],0,0,CLR_NONE,CLR_NONE,ILD_NORMAL);

		// draw text
		COLORREF color;
		if (item.id==MENU_EMPTY || item.id==MENU_EMPTY_TOP)
			color=textColors[index][bHot?3:2];
		else
			color=textColors[index][bHot?1:0];
		RECT rc={itemRect.left+iconSize+iconPadding[index].left+iconPadding[index].right+textPadding[index].left,itemRect.top+m_TextTopOffset[index],
		         itemRect.right-arrSize[index].cx-arrPadding[index].cx-arrPadding[index].cy-textPadding[index].right,itemRect.bottom-m_TextTopOffset[index]};
		DWORD flags=DT_VCENTER|DT_SINGLELINE|DT_END_ELLIPSIS;
		if (item.id==MENU_NO)
			flags|=DT_NOPREFIX;
		else if (!s_bKeyboardCues)
			flags|=DT_HIDEPREFIX;
		if (s_Theme)
		{
			DTTOPTS opts={sizeof(opts),DTT_TEXTCOLOR};
			opts.crText=color;
			if (glow[index] || opacity[index]==MenuSkin::OPACITY_FULLALPHA || opacity[index]==MenuSkin::OPACITY_FULLGLASS)
				opts.dwFlags|=DTT_COMPOSITED;
			if (glow[index])
			{
				opts.dwFlags|=DTT_GLOWSIZE;
				opts.iGlowSize=glow[index];
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
			if (bmpArrow[index])
			{
				int x=itemRect.right-arrPadding[index].cy-arrSize[index].cx;
				int y=(itemRect.top+itemRect.bottom-arrSize[index].cy)/2;
				HBITMAP bmp0=(HBITMAP)SelectObject(hdc2,bmpArrow[index]);
				if (bArr32[index])
				{
					BLENDFUNCTION func={AC_SRC_OVER,0,255,AC_SRC_ALPHA};
					AlphaBlend(hdc,x,y,arrSize[index].cx,arrSize[index].cy,hdc2,0,bHot?arrSize[index].cy:0,arrSize[index].cx,arrSize[index].cy,func);
				}
				else
				{
					BitBlt(hdc,x,y,arrSize[index].cx,arrSize[index].cy,hdc2,0,bHot?arrSize[index].cy:0,SRCCOPY);
				}
				SelectObject(hdc2,bmp0);
			}
			else
			{
				int x=itemRect.right-arrPadding[index].cy-4;
				int y=(itemRect.top+itemRect.bottom-6)/2;
				HBITMAP bmp0=(HBITMAP)SelectObject(hdc2,m_ArrowsBitmap[index*2+(bHot?1:0)]);
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
			HBITMAP bmp0=(HBITMAP)SelectObject(hdc2,m_ArrowsBitmap[0]); // the insert mask can't be in the second column of the main menu
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
	MenuSkin::TOpacity opacity=(m_bSubMenu?s_Skin.Submenu_opacity:s_Skin.Main_opacity);
	if ((opacity==MenuSkin::OPACITY_GLASS || opacity==MenuSkin::OPACITY_FULLGLASS) && ((!m_bSubMenu && m_Bitmap) || (m_bSubMenu && s_Skin.Submenu_bitmap)))
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
		if (m_bTwoColumns && (s_Skin.Main_opacity2==MenuSkin::OPACITY_GLASS || s_Skin.Main_opacity2==MenuSkin::OPACITY_ALPHA))
		{
			RECT rc;
			IntersectRect(&rc,&ps.rcPaint,&m_rContent2);
			BufferedPaintSetAlpha(hBufferedPaint,&rc,255);
		}
		EndBufferedPaint(hBufferedPaint,TRUE);
	}

	if (uMsg!=WM_PRINTCLIENT)
		EndPaint(&ps);

	return 0;
}
