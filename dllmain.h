// Copyright (C) 2014 Chris Richardson
// Please see the file file license.txt included with this project for licensing.
// If no license.txt was included, see the following URL: http://opensource.org/licenses/MIT

// dllmain.h : Declaration of module class.

class CStatusBarExtModule : public ATL::CAtlDllModuleT< CStatusBarExtModule >
{
public :
	DECLARE_LIBID(LIBID_StatusBarExtLib)
	DECLARE_REGISTRY_APPID_RESOURCEID(IDR_STATUSBAREXT, "{C4478CE9-1C4E-47E8-B738-D97BEA6FDC61}")
};

extern class CStatusBarExtModule _AtlModule;
