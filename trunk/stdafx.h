// Copyright (C) 2014 Chris Richardson
// Please see the file file license.txt included with this project for licensing.
// If no license.txt was included, see the following URL: http://opensource.org/licenses/MIT

// stdafx.h : include file for standard system include files,
// or project specific include files that are used frequently,
// but are changed infrequently

#pragma once

#ifndef STRICT
#define STRICT
#endif

#define WIN32_LEAN_AND_MEAN

#include "targetver.h"

#define _ATL_APARTMENT_THREADED

#define _ATL_NO_AUTOMATIC_NAMESPACE

#define _ATL_CSTRING_EXPLICIT_CONSTRUCTORS	// some CString constructors will be explicit


#define ATL_NO_ASSERT_ON_DESTROY_NONEXISTENT_WINDOW

#include "resource.h"
#include <atlbase.h>
#include <atlcom.h>
#include <atlctl.h>

// Shell headers
#include <Shobjidl.h>
#include <Shlguid.h>
#include <propkey.h>
#include <propvarutil.h>
#include <propsys.h>
#include <exdispid.h>
