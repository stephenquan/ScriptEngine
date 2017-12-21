#pragma once

class CGlobals : public IDispatch
{
public:
	CGlobals();
	~CGlobals();

	// IUnknown

	STDMETHOD(QueryInterface)(REFIID riid, void ** ppvObject);
	STDMETHOD_(ULONG, AddRef)();
	STDMETHOD_(ULONG, Release)();

	// IDispatch
	STDMETHOD(GetTypeInfoCount)(UINT *pctinfo);
	STDMETHOD(GetTypeInfo)(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo);
	STDMETHOD(GetIDsOfNames)(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId);
	STDMETHOD(Invoke)( DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr);

	// Misc
	STDMETHOD(Clear)();
	STDMETHOD(AddItem)(const VARIANT& item);

	static HRESULT GetDispatch(const VARIANT& item, IDispatch** ppIDispatch);

protected:
	LONG m_Ref;
	CComSafeArray<VARIANT> m_Items;

};
