// stdafx.h : include file for standard system include files,
// or project specific include files that are used frequently, but
// are changed infrequently
//

#pragma once

#include "targetver.h"

#define STRICT_TYPED_ITEMIDS
// Windows Header Files:
#include <windows.h>
#include <commctrl.h>
#include <shlobj.h>
#include <shellapi.h>

#define _ATL_CSTRING_EXPLICIT_CONSTRUCTORS      // some CString constructors will be explicit

#include <atlbase.h>
#include <atlstr.h>
#include <atlwin.h>

extern HINSTANCE g_Instance;

#ifdef BUILD_SETUP
#define INI_PATH L""
#define DOC_PATH L""
#else
#define INI_PATH L"..\\"
#define DOC_PATH L"..\\..\\Docs\\"
#endif

#include "StringUtils.h"
