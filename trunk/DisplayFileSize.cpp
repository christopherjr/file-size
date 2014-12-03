// Copyright (C) 2014 Chris Richardson
// Please see the file file license.txt included with this project for licensing.
// If no license.txt was included, see the following URL: http://opensource.org/licenses/MIT

// DisplayFileSize.cpp : Implementation of CDisplayFileSize

#include "stdafx.h"
#include "DisplayFileSize.h"


// CDisplayFileSize

#include <Shobjidl.h>
#include <Shlguid.h>
#include <shdispid.h>
#include <propkey.h>
#include <propvarutil.h>

// ToDo:
//
// Try to subclass the main view to handle the ESCAPE key.  This could then
//  perform a deselect-all operation, which I would find useful.
//
// The file size calculation and update code is running each time we activate
//  an explorer window, even if the update is not necessary. We should put in
//  some filtering so it only runs when the size actually needs to be updated.
//
// Add disk free space display for appropriate folders (not sure how this is
//  done on XP).

// Add file modified time display if a single file is selected.
//
// We may want to make the trigger timer a bit smarter.  It is currently always
// expiring in 100ms, which is OK for cases I tested.  It might be better to
// make it grow slightly if we notice that a lot of the size calculations are
// taking a long time.

// Change this to 1 if you want to see some logging in the debugger.
#define DEBUG_LOGGING 0

// Timer ID for disk space status pane update.
#define IDT_UPDATE_FILE_SIZE_PANE 0xf11e51fe

// Find the status bar window given a handle to the main explorer window.
// Obviously not optimal.
// We could possibly interrogate the browser for the status bar instead.
HWND FindStatusBar( HWND p_hExplorerWnd )
{
   HWND a_hResult = NULL;
   HWND a_hChild = GetWindow( p_hExplorerWnd, GW_CHILD );
   while( a_hChild )
   {
      TCHAR a_atcClassName[128] = {0};
      GetClassName( a_hChild, a_atcClassName, _countof( a_atcClassName ) );
      if( !_tcscmp( a_atcClassName, _T("msctls_statusbar32") ) )
      {
         a_hResult = a_hChild;
         break;
      }
      else
      {
         // Search children.
         a_hResult = FindStatusBar( a_hChild );
         if( a_hResult )
         {
            break;
         }
      }
      a_hChild = GetWindow( a_hChild, GW_HWNDNEXT );
   }

   return a_hResult;
}

// CDisplayFileSize

CDisplayFileSize::CDisplayFileSize() :
   c_hExplorerWnd( NULL ),
   c_hDefView( NULL ),
   c_hStatusBar( NULL ),
   c_iPartIndex( 0 ),
   c_bUseSbSetTextSelectionChange( false )
{
   c_oWebBrowser2Sink.Attach(new CWebBrowser2EventsSink( this ));
   c_oShellFolderViewSink.Attach(new CShellFolderViewEventsSink( this ));
}

CDisplayFileSize::~CDisplayFileSize()
{
}

HRESULT CDisplayFileSize::SetSite(IUnknown *pUnkSite)
{
   HRESULT a_hResult = S_OK;

   // AddRef the new site before releasing any existing site, so we don't run
   // into problems if the site is the same as our existing site.
   if( pUnkSite )
   {
      pUnkSite->AddRef();
   }

   if( c_hDefView )
   {
      RemoveWindowSubclass( c_hDefView, DefViewProcS, 0 );
   }

   // Detach from the status bar window so we won't try to update it while we
   // are in the middle of this work.
   if( c_hStatusBar )
   {
      RemoveWindowSubclass( c_hStatusBar, StatusBarProcS, 0 );
   }

   // Now detach from the existing site and other interfaces.
   // Disconnect event handlers and release other interfaces retrieved from the
   // site before we release the site.
   if( c_oWebBrowser2Sink )
   {
      c_oWebBrowser2Sink->Disconnect();
   }

   if( c_oShellFolderViewSink )
   {
      c_oShellFolderViewSink->Disconnect();
   }

   c_oWebBrowser.Release();
   c_oShellBrowser.Release();
   c_oServiceProvider.Release();
   c_oPropertySystem.Release();
   c_oFolderView.Release();

   c_oSite.Release();

   c_oSite.Attach( pUnkSite );
   c_hExplorerWnd = NULL;
   c_hDefView = NULL;
   c_hStatusBar = NULL;

   a_hResult = IObjectWithSiteImpl<CDisplayFileSize>::SetSite(pUnkSite);
   if( FAILED( a_hResult ) )
   {
      return a_hResult;
   }

   // Now hook everything up using the new site.
   if( c_oSite )
   {
      // The service provider lets us look up the IFolderView2 later, and the
      // property system lets us format the file size display strings.
      c_oServiceProvider = pUnkSite;
      PSGetPropertySystem( IID_PPV_ARGS( &c_oPropertySystem ) );

      // The IWebBrowser2 will give us the main window handle.
      // We might also be able to use its status bar features directly but I
      // haven't investigated that.
      CComQIPtr<IWebBrowser2> a_oBrowser( pUnkSite );
      if( a_oBrowser )
      {
         SHANDLE_PTR a_pHWnd = NULL;
         a_hResult = a_oBrowser->get_HWND( &a_pHWnd );

         if( SUCCEEDED( a_hResult ) && a_pHWnd )
         {
            c_hExplorerWnd = reinterpret_cast<HWND>(a_pHWnd);
         }

         // Register for events now that we are all hooked up to the UI.
         // Not much we can do if this fails.
         c_oWebBrowser = a_oBrowser;
         c_oWebBrowser2Sink->Connect( c_oWebBrowser );
      }
   }
   return S_OK;
}

HRESULT CDisplayFileSize::OnWebBrowser2EventsInvoke( DISPID p_eDispId, DISPPARAMS * p_poDispParams, VARIANT * p_poVarResult )
{
   HRESULT a_hResult = S_OK;

   UNREFERENCED_PARAMETER( p_poDispParams );
   UNREFERENCED_PARAMETER( p_poVarResult );

   switch( p_eDispId )
   {
      case DISPID_DOCUMENTCOMPLETE:
         a_hResult = OnWebBrowser2DocumentComplete();
         break;
   }

   return a_hResult;
}

HRESULT CDisplayFileSize::OnWebBrowser2DocumentComplete()
{
   HRESULT a_hResult = S_OK;
   if( c_oServiceProvider )
   {
      IFolderView2 * a_poFolderView = NULL;

      // Try to get the new FolderView, then release our existing FolderView and
      // attach to the new FolderView if we got it.
      a_hResult = c_oServiceProvider->QueryService( SID_SFolderView,
                                                    IID_PPV_ARGS( &a_poFolderView ) );

      c_oFolderView.Release();
      if( SUCCEEDED( a_hResult ) )
      {
         // Note that attach does not AddRef, which is what we want.
         c_oFolderView.Attach( a_poFolderView );
      }

      c_oShellBrowser.Release();
      a_hResult = c_oServiceProvider->QueryService( SID_STopLevelBrowser,
                                                    IID_PPV_ARGS( &c_oShellBrowser ) );

      // Try to get the status bar from the IShellBrowser first.  If we can't
      // find it that way, search the window hierarchy.
      if( c_oShellBrowser )
      {
         a_hResult = c_oShellBrowser->GetControlWindow( FCW_STATUS, &c_hStatusBar );
         if( FAILED( a_hResult ) )
         {
            c_hStatusBar = NULL;
         }
      }

      if( !c_hStatusBar )
      {
         // OK, try to find the status bar via the window hierarchy.
         c_hStatusBar = FindStatusBar( c_hExplorerWnd );
      }

      if( c_hStatusBar )
      {
         // However we managed to find it, we will subclass it now so we can
         // add back the size pane.
         SetWindowSubclass( c_hStatusBar, StatusBarProcS, 0, reinterpret_cast<DWORD_PTR>(this) );
      }

      c_oShellFolderViewSink->Disconnect();
      c_bUseSbSetTextSelectionChange = true;
      if( c_oWebBrowser )
      {
         // How to get the right object to connect DShellFolderViewEvents.
         // https://social.msdn.microsoft.com/Forums/ie/en-US/e558984e-d239-4dc1-b7e0-77041ba7c34e/sinking-dshellfolderviewevents-in-c-atl-bho-sink-not-firing?forum=ieextensiondevelopment

         CComPtr<IDispatch> a_poDisp;
         a_hResult = c_oWebBrowser->get_Document( &a_poDisp );
         if( SUCCEEDED( a_hResult ) )
         {
            CComQIPtr<IShellFolderViewDual> a_oShellFolder( a_poDisp );
            if( a_oShellFolder )
            {
               a_hResult = c_oShellFolderViewSink->Connect( a_oShellFolder );
               if( SUCCEEDED( a_hResult ) )
               {
                  // We can use shell folder view events to handle selection
                  // changes rather than the SB_SETTEXT message approach.
                  c_bUseSbSetTextSelectionChange = false;
               }
            }
         }
      }

      // Subclass the shell view so we can intercept some keys for navigation.
      // We want to query for IID_CDefView before attempting to subclass the
      // DirectUIHWnd child of the DefView, to make sure we only subclass the
      // default view and don't interfere with custom namespace views.
      // http://msdn.microsoft.com/en-us/library/windows/desktop/bb774834.aspx

      CComPtr<IShellView> a_oShellView;
      a_hResult = c_oShellBrowser->QueryActiveShellView( &a_oShellView );
      if( SUCCEEDED( a_hResult ) && a_oShellView )
      {
         CComQIPtr<IShellView, &IID_CDefView> a_oDefViewQuery( a_oShellView );
         if( a_oDefViewQuery )
         {
            a_hResult = a_oShellView->GetWindow( &c_hDefView );
            if( SUCCEEDED( a_hResult ) )
            {
               c_hDefView = GetWindow( c_hDefView, GW_CHILD );
               SetWindowSubclass( c_hDefView, DefViewProcS, 0, reinterpret_cast<DWORD_PTR>(this) );
            }
         }
      }
   }

   return a_hResult;
}

HRESULT CDisplayFileSize::OnShellFolderEventsInvoke( DISPID p_eDispId, DISPPARAMS * p_poDispParams, VARIANT * p_poVarResult )
{
   HRESULT a_hResult = S_OK;

   UNREFERENCED_PARAMETER( p_poDispParams );
   UNREFERENCED_PARAMETER( p_poVarResult );
   
   switch( p_eDispId ) {
      case DISPID_SELECTIONCHANGED:
         a_hResult = OnShellFolderViewSelectionChanged();
         break;
   }

   return a_hResult;
}

HRESULT CDisplayFileSize::OnShellFolderViewSelectionChanged()
{
   #if DEBUG_LOGGING
   OutputDebugString( _T("Selection changed\n") );
   #endif

   // Trigger a timed update rather than updating immediately.
   TriggerFileSizePaneUpdate();

   return S_OK;
}

void CDisplayFileSize::TriggerFileSizePaneUpdate()
{
   // Kill and restart a 100ms timer.  If it expires before the user performs
   // another action that triggers it, then we will perform an update.
   // This prevents us from doing a huge amount of updates as the user performs
   // a lot of actions in a row, such as using the keys to scroll through a
   // large folder.
   KillTimer( c_hStatusBar, IDT_UPDATE_FILE_SIZE_PANE );
   SetTimer( c_hStatusBar, IDT_UPDATE_FILE_SIZE_PANE, 100, NULL );
}

void CDisplayFileSize::DoFileSizePaneUpdate()
{
   HRESULT a_hResult = S_OK;
   TCHAR a_atcText[MAX_TEXT] = { 0 };

   if( c_oFolderView )
   {
      CComPtr<IShellItemArray> a_oSelection;
      CComPtr<IPropertyStore> a_oPropStore;
      DWORD a_dwSelCount = 0;
      int a_iItemCount = 0;
      bool a_bHaveValue = false;
      PROPVARIANT a_vtTotalSize;
      PropVariantInit( &a_vtTotalSize );

      // We could pass in true here to treat an empty selection as representing
      // the whole folder, but getting the size property for that returns VT_EMPTY
      // which we can't do much with.  It also wouldn't handle cases where the
      // property store fails to give us a size value.
      // So, if there is no selection just retrieve all the items in the folder
      // and proceed as normal.
      a_hResult = c_oFolderView->GetSelection( FALSE, &a_oSelection );
      if( FAILED( a_hResult ) )
      {
         if( a_hResult == S_FALSE ||
             a_hResult == HRESULT_FROM_WIN32( ERROR_NOT_FOUND ) )
         {
            // No selection; no big deal.  Load up the whole folder.
            a_hResult = c_oFolderView->Items( SVGIO_ALLVIEW, IID_PPV_ARGS( &a_oSelection ) );
         }

         if( FAILED( a_hResult ) )
         {
            a_oSelection.Detach();
            SetFileSizePaneText( a_atcText );
            return;
         }
      }

      a_hResult = a_oSelection->GetCount( &a_dwSelCount );
      if( FAILED( a_hResult ) )
      {
         a_dwSelCount = 0;
      }

      // If we have a large number of selected items, update the status text
      // before we ask for the total size, so the user can at least see we are
      // doing something.
      if( a_dwSelCount >= FILE_SIZE_UPDATING_THRESHOLD )
      {
         SetFileSizePaneText( _T("Updating ...") );
      }

      //a_hResult = a_oSelection->GetPropertyStore( GPS_DEFAULT, IID_PPV_ARGS( &a_oPropStore ) );
      a_hResult = a_oSelection->GetPropertyStore( GPS_FASTPROPERTIESONLY, IID_PPV_ARGS( &a_oPropStore ) );
      if( FAILED( a_hResult ) )
      {
         a_oPropStore.Detach();
      }

      if( a_oPropStore )
      {
         a_hResult = a_oPropStore->GetValue( PKEY_Size, &a_vtTotalSize );
         if( SUCCEEDED( a_hResult ) )
         {
            a_bHaveValue = true;
         }
      }

      // GetPropertyStore can fail sometimes if there are invalid
      // (or incorrectly formatted) files selected, so we have a fallback where
      // we just iterate the selection manually.
      if( !a_bHaveValue )
      {
         // Accumulate the size of each item.
         ULONGLONG a_ullTotal = 0;
         ULONGLONG a_ullTemp = 0;
         for( DWORD i = 0; i < a_dwSelCount; ++i )
         {
            CComPtr<IShellItem> a_oShellItem;
            a_hResult = a_oSelection->GetItemAt( i, &a_oShellItem );
            if( SUCCEEDED( a_hResult ) && a_oShellItem )
            {
               CComQIPtr<IShellItem2> a_oShellItem2( a_oShellItem );
               if( a_oShellItem2 )
               {
                  a_hResult = a_oShellItem2->GetUInt64( PKEY_Size, &a_ullTemp );
                  if( SUCCEEDED( a_hResult ) )
                  {
                     a_ullTotal += a_ullTemp;
                  }
               }
            }
         }

         a_hResult = InitPropVariantFromUInt64( a_ullTotal, &a_vtTotalSize );
         if( FAILED( a_hResult ) )
         {
            PropVariantInit( &a_vtTotalSize );
         }
      }

      // Format the display string, by asking the property system if we have
      // a valid IPropertySystem, or manually otherwise.
      if( c_oPropertySystem )
      {
         c_oPropertySystem->FormatForDisplay( PKEY_Size,
                                              a_vtTotalSize,
                                              PDFF_DEFAULT,
                                              a_atcText,
                                              _countof( a_atcText ) );
      }
      else
      {
         // FIXME
         // Use some other method besides the IPropertySystem to format the
         // value. I don't really feel like writing that.
         ULONGLONG a_ullTotalSize = 0;
         a_hResult = PropVariantToUInt64( a_vtTotalSize, &a_ullTotalSize );
         if( FAILED( a_hResult ) )
         {
            a_ullTotalSize = 0;
         }
         // PRIu64 expands into two narrow literals so we can't use it for now.
         // Use %llu explicitly.
         // http://stackoverflow.com/a/21789920
         _sntprintf_s( a_atcText, _TRUNCATE, _T("%llu"), a_ullTotalSize );
      }

      PropVariantClear( &a_vtTotalSize );
   }

   SetFileSizePaneText( a_atcText );
}

void CDisplayFileSize::SetFileSizePaneText( const TCHAR * p_ptcText )
{
   SendMessage( c_hStatusBar, SB_SETTEXTW, MAKEWPARAM( MAKEWORD( c_iPartIndex, SBT_POPOUT ), 0 ), reinterpret_cast<LPARAM>(p_ptcText) );
}

LRESULT CALLBACK CDisplayFileSize::StatusBarProcS( HWND hWnd,
                                                   UINT uMsg,
                                                   WPARAM wParam,
                                                   LPARAM lParam,
                                                   UINT_PTR uIdSubclass,
                                                   DWORD_PTR dwRefData )
{
   return reinterpret_cast<CDisplayFileSize*>(dwRefData)->StatusBarProc( hWnd, uMsg, wParam, lParam, uIdSubclass );
}

LRESULT CALLBACK CDisplayFileSize::StatusBarProc( HWND hWnd,
                                                  UINT uMsg,
                                                  WPARAM wParam,
                                                  LPARAM lParam,
                                                  UINT_PTR uIdSubclass )
{
   LRESULT a_lResult = 0;
   bool a_bCallDefProc = true;

   UNREFERENCED_PARAMETER( uIdSubclass );

   switch( uMsg )
   {
      case WM_NCDESTROY:
         // Per Raymond Chen (http://blogs.msdn.com/oldnewthing/archive/2003/11/11/55653.aspx),
         // we must remove our subclass when the window is destroyed, if we haven't already done so.
         OnSbWmNcDestroy();
         break;

      case WM_TIMER:
         OnSbWmTimer( static_cast<UINT_PTR>(wParam) );
         break;

      case SB_SETPARTS:
         a_lResult = OnSbSetParts( static_cast<int>(wParam), reinterpret_cast<int*>(lParam) );
         a_bCallDefProc = false;
         break;

      case SB_SETTEXTW:
         OnSbSetTextW( static_cast<int>(LOBYTE( LOWORD( wParam ) )), static_cast<int>(HIBYTE( LOWORD( wParam ) )), reinterpret_cast<const WCHAR*>(lParam) );
         break;

      default:
      {
         // For debugging.
         #if DEBUG_LOGGING
         TCHAR a_atcTemp[MAX_TEXT] = {0};
         _sntprintf_s( a_atcTemp, _TRUNCATE, _T("Msg: 0x%08X (%d)  0x%08X 0x%08X\n"), uMsg, uMsg, wParam, lParam );
         OutputDebugString( a_atcTemp );
         #endif
         break;
      }
   }

   if( a_bCallDefProc )
   {
      a_lResult = DefSubclassProc( hWnd, uMsg, wParam, lParam );
   }
   return a_lResult;
}

void CDisplayFileSize::OnSbWmNcDestroy()
{
   RemoveWindowSubclass( c_hStatusBar, StatusBarProcS, 0 );
}

void CDisplayFileSize::OnSbWmTimer( UINT_PTR p_uiTimerId )
{
   KillTimer( c_hStatusBar, p_uiTimerId );
   switch( p_uiTimerId )
   {
      case IDT_UPDATE_FILE_SIZE_PANE:
         DoFileSizePaneUpdate();
         break;
   }
}

LRESULT CDisplayFileSize::OnSbSetParts( int p_iPartCount, int * p_piParts )
{
   #if DEBUG_LOGGING
   TCHAR a_atcTemp[MAX_TEXT] = {0};
   _sntprintf_s( a_atcTemp, _TRUNCATE, _T("SetParts: %d\n"), p_iPartCount );
   OutputDebugString( a_atcTemp );
   #endif

   int a_iPartCount = 0;
   int a_aiParts[MAX_STATUS_BAR_PARTS] = {0};
   RECT a_stClient = {0};
   GetClientRect( c_hStatusBar, &a_stClient );

   c_iPartIndex = 0;

   // Make sure the part count is valid. Note we subtract 1 from the max parts
   // allowed since we will insert our own part below if necessary.
   if( p_iPartCount < 0 || p_iPartCount > (MAX_STATUS_BAR_PARTS - 1) )
   {
      return FALSE;
   }

   // Insert the maximum of the requested number of parts or 3, so we always
   // have room for the status help text, the file size and the zone icon
   // ("My Computer", "Internet", etc).
   a_iPartCount = max( p_iPartCount, MINIMUM_PART_COUNT );
   c_iPartIndex = 1;
   if( p_piParts && p_iPartCount >= MINIMUM_PART_COUNT )
   {
      // Use the requested part layout directly and just overwrite whatever
      // our caller was going to place at index 1 when we set the text later.
      memcpy( a_aiParts, p_piParts, p_iPartCount * sizeof( a_aiParts[0] ) );
   }
   else
   {
      // There are no requested parts. Just insert a dummy part to the left of
      // our part, so our part is near the right edge rather than the left edge.
      // The size of these parts may not consistent with the sizes explorer uses,
      // but they work OK for me.
      a_aiParts[0] = a_stClient.right - MY_COMPUTER_SIZE_PART_WIDTH - FILE_SIZE_PART_WIDTH;
      a_aiParts[1] = a_stClient.right - MY_COMPUTER_SIZE_PART_WIDTH;
      a_aiParts[2] = -1;
   }

   return DefSubclassProc( c_hStatusBar, SB_SETPARTS, static_cast<WPARAM>(a_iPartCount), reinterpret_cast<LPARAM>(a_aiParts) );
}

LRESULT CDisplayFileSize::OnSbSetTextW( int p_iPartIndex, int p_iDrawOp, const WCHAR * p_pwcText )
{
   #if DEBUG_LOGGING
   TCHAR a_atcTemp[MAX_TEXT] = {0};
   _sntprintf_s( a_atcTemp, _TRUNCATE, _T("SetText: %d %d %s\n"), p_iPartIndex, p_iDrawOp, p_pwcText );
   OutputDebugString( a_atcTemp );
   #else
   UNREFERENCED_PARAMETER( p_pwcText );
   UNREFERENCED_PARAMETER( p_iDrawOp );
   #endif

   // This message is used to trigger an update to the status bar only if we
   // were not able to successfully sink DWebBrowserEvents2 in order to be
   // properly notified of selection changes from the explorer browser.
   if( c_bUseSbSetTextSelectionChange )
   {
      // The default Shell View seems to send this message as the selection is
      // changed, which gives us a reasonable hook to use to trigger updates to
      // our status pane.
      // This is not optimal, since this message is sent in other cases as well,
      // where we don't need to trigger an update.  For now it works well enough
      // for me, but any improvements are welcome.

      // Only trigger an update if the part being updated is not *our* part, since
      // we don't want to trigger more updates as we change the text on our part.
      if( p_iPartIndex != c_iPartIndex )
      {
         TriggerFileSizePaneUpdate();
      }
   }

   return TRUE;
}



LRESULT CALLBACK CDisplayFileSize::DefViewProcS( HWND hWnd,
                                                 UINT uMsg,
                                                 WPARAM wParam,
                                                 LPARAM lParam,
                                                 UINT_PTR uIdSubclass,
                                                 DWORD_PTR dwRefData )
{
   return reinterpret_cast<CDisplayFileSize*>(dwRefData)->DefViewProc( hWnd, uMsg, wParam, lParam, uIdSubclass );
}

LRESULT CALLBACK CDisplayFileSize::DefViewProc( HWND hWnd,
                                                UINT uMsg,
                                                WPARAM wParam,
                                                LPARAM lParam,
                                                UINT_PTR uIdSubclass )
{
   LRESULT a_lResult = 0;
   bool a_bCallDefProc = true;

   UNREFERENCED_PARAMETER( uIdSubclass );

   switch( uMsg )
   {
      case WM_NCDESTROY:
         OnDvWmNcDestroy();
         break;

      case WM_KEYDOWN:
         if( OnDvWmKeyDown( static_cast<UINT>(wParam), static_cast<UINT>(lParam) ) )
         {
            a_bCallDefProc = false;
         }
         break;

      default:
      {
         // For debugging.
         #if DEBUG_LOGGING
         TCHAR a_atcTemp[MAX_TEXT] = {0};
         _sntprintf_s( a_atcTemp, _TRUNCATE, _T("Msg: 0x%08X (%d)  0x%08X 0x%08X\n"), uMsg, uMsg, wParam, lParam );
         OutputDebugString( a_atcTemp );
         #endif
         break;
      }
   }

   if( a_bCallDefProc )
   {
      a_lResult = DefSubclassProc( hWnd, uMsg, wParam, lParam );
   }
   return a_lResult;
}

void CDisplayFileSize::OnDvWmNcDestroy()
{
   RemoveWindowSubclass( c_hDefView, DefViewProcS, 0 );
}

bool CDisplayFileSize::OnDvWmKeyDown( UINT p_uiVK, UINT p_uiFlags )
{
   bool a_bHandled = false;

   UNREFERENCED_PARAMETER( p_uiFlags );

   switch( p_uiVK )
   {
      case 'U':
         if( (GetKeyState( VK_CONTROL ) & 0x8000000) != 0 )
         {
            if( !c_oShellBrowser )
            {
               break;
            }

            // Navigate up one level.
            c_oShellBrowser->BrowseObject( NULL, SBSP_SAMEBROWSER | SBSP_PARENT );
            a_bHandled = true;
         }
         break;

      case VK_ESCAPE:
      {
         if( !c_oFolderView )
         {
            break;
         }

         // If there are any selected items, deselect them all.
         // Otherwise, trigger an update to show the size of all the
         // items in the folder.  This allows the user to show the size
         // of a folder when they first navigate to it, by clicking Escape.
         int a_iSelectedCount = 0;
         c_oFolderView->ItemCount( SVGIO_SELECTION, &a_iSelectedCount );
         if( a_iSelectedCount )
         {
            c_oFolderView->SelectItem( -1, SVSI_DESELECTOTHERS );
            a_bHandled = true;
         }
         else
         {
            TriggerFileSizePaneUpdate();
         }

         // This is nice if we want to go back to the beginning of the
         // file list after we have already selected a file in the view.
         // Then we don't have to use the Home key in order to perform a
         // new keyboard search in the view.
         c_oFolderView->SelectItem( 0, SVSI_FOCUSED );
         break;
      }
   }

   return a_bHandled;
}


// CDisplayFileSize::CWebBrowser2EventsSink

CDisplayFileSize::CWebBrowser2EventsSink::CWebBrowser2EventsSink( CDisplayFileSize * p_poOuter ) :
   c_poOuter( p_poOuter )
{
}

HRESULT CDisplayFileSize::CWebBrowser2EventsSink::SimpleInvoke(
   DISPID dispid, DISPPARAMS * pdispparams, VARIANT * pvarResult )
{
   return c_poOuter->OnWebBrowser2EventsInvoke( dispid, pdispparams, pvarResult );
}


// CDisplayFileSize::CShellFolderViewEventsSink

CDisplayFileSize::CShellFolderViewEventsSink::CShellFolderViewEventsSink( CDisplayFileSize * p_poOuter ) :
   c_poOuter( p_poOuter )
{
}

HRESULT CDisplayFileSize::CShellFolderViewEventsSink::SimpleInvoke(
   DISPID dispid, DISPPARAMS * pdispparams, VARIANT * pvarResult )
{
   return c_poOuter->OnShellFolderEventsInvoke( dispid, pdispparams, pvarResult );
}

