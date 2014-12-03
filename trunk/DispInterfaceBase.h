// The code for CDispInterfaceBase is by Raymond Chen, from his weblog at
// http://blogs.msdn.com/b/oldnewthing/archive/2013/06/12/10425215.aspx.
// The license is not explicit.
//
// Changes from the original code are licensed as follows:
// Copyright (C) 2014 Chris Richardson
// Please see the file file license.txt included with this project for licensing.
// If no license.txt was included, see the following URL: http://opensource.org/licenses/MIT

// CDispInterfaceBase.h : Declaration of the CDispInterfaceBase

#pragma once

template<typename DispInterface>
class CDispInterfaceBase : public DispInterface
{
public:
   CDispInterfaceBase() : m_cRef( 1 ), m_dwCookie( 0 ) { }

   /* IUnknown */
   IFACEMETHODIMP QueryInterface( REFIID riid, void **ppv )
   {
      *ppv = nullptr;
      HRESULT hr = E_NOINTERFACE;
      if( riid == IID_IUnknown || riid == IID_IDispatch ||
          riid == __uuidof(DispInterface) )
      {
         *ppv = static_cast<DispInterface *>
            (static_cast<IDispatch*>(this));
         AddRef();
         hr = S_OK;
      }
      return hr;
   }

   IFACEMETHODIMP_( ULONG ) AddRef()
   {
      return InterlockedIncrement( &m_cRef );
   }

   IFACEMETHODIMP_( ULONG ) Release()
   {
      LONG cRef = InterlockedDecrement( &m_cRef );
      if( cRef == 0 ) delete this;
      return cRef;
   }

   // *** IDispatch ***
   IFACEMETHODIMP GetTypeInfoCount( UINT *pctinfo )
   {
      *pctinfo = 0;
      return E_NOTIMPL;
   }

   IFACEMETHODIMP GetTypeInfo( UINT /*iTInfo*/, LCID /*lcid*/,
                               ITypeInfo **ppTInfo )
   {
      *ppTInfo = nullptr;
      return E_NOTIMPL;
   }

   IFACEMETHODIMP GetIDsOfNames( REFIID, LPOLESTR * /*rgszNames*/,
                                 UINT /*cNames*/, LCID /*lcid*/,
                                 DISPID * /*rgDispId*/ )
   {
      return E_NOTIMPL;
   }

   IFACEMETHODIMP Invoke(
      DISPID dispid, REFIID /*riid*/, LCID /*lcid*/, WORD /*wFlags*/,
      DISPPARAMS *pdispparams, VARIANT *pvarResult,
      EXCEPINFO * /*pexcepinfo*/, UINT * /*puArgErr*/ )
   {
      if( pvarResult ) VariantInit( pvarResult );
      return SimpleInvoke( dispid, pdispparams, pvarResult );
   }

   // Derived class must implement SimpleInvoke
   virtual HRESULT SimpleInvoke( DISPID dispid,
                                 DISPPARAMS *pdispparams, VARIANT *pvarResult ) = 0;

public:
   HRESULT Connect( IUnknown *punk )
   {
      HRESULT hr = S_OK;
      CComPtr<IConnectionPointContainer> spcpc;
      if( SUCCEEDED( hr ) ) {
         hr = punk->QueryInterface( IID_PPV_ARGS( &spcpc ) );
      }
      if( SUCCEEDED( hr ) ) {
         hr = spcpc->FindConnectionPoint( __uuidof(DispInterface), &m_spcp );
      }
      if( SUCCEEDED( hr ) ) {
         hr = m_spcp->Advise( this, &m_dwCookie );
      }
      return hr;
   }

   void Disconnect()
   {
      if( m_dwCookie ) {
         m_spcp->Unadvise( m_dwCookie );
         m_spcp.Release();
         m_dwCookie = 0;
      }
   }

private:
   LONG m_cRef;
   ATL::CComPtr<IConnectionPoint> m_spcp;
   DWORD m_dwCookie;
};
