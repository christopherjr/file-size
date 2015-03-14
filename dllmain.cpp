// Copyright (C) 2015 Chris Richardson
// Please see the file file license.txt included with this project for licensing.
// If no license.txt was included, see the following URL: http://opensource.org/licenses/MIT

// dllmain.cpp : Implementation of DllMain.

#include "stdafx.h"
#include "resource.h"
#include "StatusBarExt_i.h"
#include "dllmain.h"

CStatusBarExtModule _AtlModule;

// DLL Entry Point
extern "C" BOOL WINAPI DllMain(HINSTANCE hInstance, DWORD dwReason, LPVOID lpReserved)
{
   if( dwReason == DLL_PROCESS_ATTACH )
   {
      // Don't load in IE.
      TCHAR pszLoader[MAX_PATH];
      GetModuleFileName(NULL, pszLoader, MAX_PATH);
      _tcslwr_s(pszLoader);
      if (_tcsstr(pszLoader, _T("iexplore.exe"))) 
         return FALSE;

      // We don't need to be called as threads are created and destroyed, so
      // save some resources by disabling all that.
      DisableThreadLibraryCalls( hInstance );
   }

   return _AtlModule.DllMain(dwReason, lpReserved); 
}
