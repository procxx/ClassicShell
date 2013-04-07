// Classic Shell (c) 2009-2013, Ivo Beltchev
// The sources for Classic Shell are distributed under the MIT open source license

#pragma once

#include "ClassicIE9DLL_i.h"

class CClassicIE9DLLModule : public CAtlDllModuleT< CClassicIE9DLLModule >
{
public :
	DECLARE_LIBID(LIBID_ClassicIE9DLLLib)
	DECLARE_REGISTRY_APPID_RESOURCEID(IDR_CLASSICIE9DLL, "{DF3255F4-FF55-44FA-A728-E77B83E9E403}")
};

extern class CClassicIE9DLLModule _AtlModule;
