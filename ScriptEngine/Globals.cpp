#include "stdafx.h"
#include "Globals.h"

CGlobals::CGlobals() :
    m_Ref(1)
{
	m_Items.Create();
}

CGlobals::~CGlobals()
{
}

STDMETHODIMP CGlobals::QueryInterface(REFIID riid, void ** ppvObject)
{
	if (riid == IID_IUnknown)
	{
		*(IUnknown**&) ppvObject = (IUnknown*) this;
		AddRef();
		return S_OK;
	}

	if (riid == IID_IDispatch)
	{
		*(IDispatch**&) ppvObject = (IDispatch*) this;
		AddRef();
		return S_OK;
	}

	return E_NOINTERFACE;
}

STDMETHODIMP_(ULONG) CGlobals::AddRef()
{
	InterlockedIncrement(&m_Ref);
	return m_Ref;
}

STDMETHODIMP_(ULONG) CGlobals::Release()
{
	InterlockedDecrement(&m_Ref);
	if (!m_Ref)
	{
		delete this;
		return 0;
	}
	return m_Ref;
}

STDMETHODIMP CGlobals::GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId)
{
	for (ULONG idx = 0; idx < m_Items.GetCount(); idx++)
	{
		IDispatch* pIDispatch = m_Items.GetAt(idx);
		HRESULT hr = pIDispatch->GetIDsOfNames(riid, rgszNames, cNames, lcid, rgDispId);
		if (SUCCEEDED(hr))
		{
			for (UINT iName = 0; iName < cNames; iName++)
			{
				rgDispId[iName] |= (idx << 16);
			}
			return hr;
		}
	}

	return DISP_E_UNKNOWNNAME;
}

STDMETHODIMP CGlobals::Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
{
	HRESULT hr = S_OK;
	LONG idx = (dispIdMember >> 16);
	if (idx < 0 || (ULONG&) idx >= m_Items.GetCount()) return E_INVALIDARG;
	IDispatch* pIDispatch = m_Items.GetAt(idx);
	hr = pIDispatch->Invoke(dispIdMember & 0xffff, riid, lcid, wFlags, pDispParams, pVarResult, pExcepInfo, puArgErr);
	return hr;
}

STDMETHODIMP CGlobals::GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo)
{
	return E_NOTIMPL;
}

STDMETHODIMP CGlobals::GetTypeInfoCount(UINT *pctinfo)
{
	*pctinfo = 0;
	return S_OK;
}

STDMETHODIMP CGlobals::Clear()
{
	HRESULT hr = S_OK;
	CHECKHR(m_Items.Destroy());
	CHECKHR(m_Items.Create());
	return hr;
}

STDMETHODIMP CGlobals::AddItem(IDispatch* item)
{
	HRESULT hr = S_OK;
	CHECKHR(m_Items.Add(item));
	return hr;
}
