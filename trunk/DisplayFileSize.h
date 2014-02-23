// Copyright (C) 2014 Chris Richardson
// Please see the file file license.txt included with this project for licensing.
// If no license.txt was included, see the following URL: http://opensource.org/licenses/MIT

// DisplayFileSize.h : Declaration of the CDisplayFileSize

#pragma once
#include "resource.h"       // main symbols
#include "StatusBarExt_i.h"


#if defined(_WIN32_WCE) && !defined(_CE_DCOM) && !defined(_CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA)
#error "Single-threaded COM objects are not properly supported on Windows CE platform, such as the Windows Mobile platforms that do not include full DCOM support. Define _CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA to force ATL to support creating single-thread COM object's and allow use of it's single-threaded COM object implementations. The threading model in your rgs file was set to 'Free' as that is the only threading model supported in non DCOM Windows CE platforms."
#endif

using namespace ATL;

// For handling the OnDocumentComplete event.
extern ATL::_ATL_FUNC_INFO DocumentComplete2Struct;

// CDisplayFileSize

class ATL_NO_VTABLE CDisplayFileSize :
   public CComObjectRootEx<CComSingleThreadModel>,
   public CComCoClass<CDisplayFileSize, &CLSID_DisplayFileSize>,
   public IObjectWithSiteImpl<CDisplayFileSize>,
   public IDispatchImpl<IDisplayFileSize, &IID_IDisplayFileSize, &LIBID_StatusBarExtLib, /*wMajor =*/ 1, /*wMinor =*/ 0>,
   public IDispEventImpl<1, CDisplayFileSize, &DIID_DWebBrowserEvents2, &LIBID_SHDocVw, 1, 1>
{
public:
   CDisplayFileSize();
   virtual ~CDisplayFileSize();

DECLARE_REGISTRY_RESOURCEID(IDR_DISPLAYFILESIZE)


BEGIN_COM_MAP(CDisplayFileSize)
   COM_INTERFACE_ENTRY(IDisplayFileSize)
   COM_INTERFACE_ENTRY(IDispatch)
   COM_INTERFACE_ENTRY(IObjectWithSite)
END_COM_MAP()

BEGIN_SINK_MAP(CDisplayFileSize)
    SINK_ENTRY_INFO(1, DIID_DWebBrowserEvents2, DISPID_DOCUMENTCOMPLETE, OnDocumentComplete, &DocumentComplete2Struct)
END_SINK_MAP()


   DECLARE_PROTECT_FINAL_CONSTRUCT()

   HRESULT FinalConstruct()
   {
      return S_OK;
   }

   void FinalRelease()
   {
   }

public:
   // IObjectWithSite
   STDMETHOD(SetSite)(IUnknown *pUnkSite);

   // WebBrowser2 events
   STDMETHOD(OnDocumentComplete)(IDispatch *pDisp, VARIANT *pvarURL);

private:
   // Helpers.
   void TriggerFileSizePaneUpdate();
   void DoFileSizePaneUpdate();
   void SetFileSizePaneText( const TCHAR * p_ptcText );

   // Status bar message handlers.
   LRESULT OnSbSetTextW( int p_iPartIndex, int p_iDrawOp, const WCHAR * p_pwcText );
   LRESULT OnSbSetParts( int p_iPartCount, int * p_piParts );
   void OnWmNcDestroy();
   void OnWmTimer( UINT_PTR p_uiTimerId );

   static LRESULT CALLBACK StatusBarProcS( HWND hWnd,
                                           UINT uMsg,
                                           WPARAM wParam,
                                           LPARAM lParam,
                                           UINT_PTR uIdSubclass,
                                           DWORD_PTR dwRefData );
   LRESULT CALLBACK StatusBarProc( HWND hWnd,
                                   UINT uMsg,
                                   WPARAM wParam,
                                   LPARAM lParam,
                                   UINT_PTR uIdSubclass );
private:
   // The width (in pixels) of the file size pane and MyComputer pane.
   // This is what explorer uses on my system but it will probably be different
   // for you if you use anything different than a default US-English Windows 7.
   // You could change this to calculate the size based on the font size and
   // make it more dynamic.
   enum
   {
      FILE_SIZE_PART_WIDTH          = 104,
      MY_COMPUTER_SIZE_PART_WIDTH   = 205
   };

   // We will always create this many parts.
   // For now, there is one part for the normal hint text, one part for the file
   // size, and one part for the zone.
   enum
   {
      MINIMUM_PART_COUNT            = 3
   };

   // The threshold count of selected items upon which we show a status hint
   // that we are updating the size pane.
   // This is a bit too simple and is not ideal.
   enum
   {
      FILE_SIZE_UPDATING_THRESHOLD = 100
   };

   // The status bar only supports up to 256 parts, so we can just store a
   // small fixed size array rather than allocating it dynamically.
   enum
   {
      MAX_STATUS_BAR_PARTS = 256
   };

   // Max text length for file size string and debug output strings.
   enum
   {
      MAX_TEXT = 128
   };

private:
   HWND c_hExplorerWnd;
   HWND c_hStatusBar;
   int c_iPartIndex;
   bool c_bAdvised;
   CComPtr<IUnknown> c_oSite;
   CComQIPtr<IServiceProvider> c_oServiceProvider;
   CComPtr<IPropertySystem> c_oPropertySystem;
   CComPtr<IFolderView2> c_oFolderView;
};

OBJECT_ENTRY_AUTO(__uuidof(DisplayFileSize), CDisplayFileSize)
