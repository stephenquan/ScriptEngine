// ScriptEngine.h : Declaration of the CScriptEngine

#pragma once
#include "resource.h"       // main symbols

#include "ScriptEngineLib_i.h"
#include "Globals.h"

#if defined(_WIN32_WCE) && !defined(_CE_DCOM) && !defined(_CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA)
#error "Single-threaded COM objects are not properly supported on Windows CE platform, such as the Windows Mobile platforms that do not include full DCOM support. Define _CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA to force ATL to support creating single-thread COM object's and allow use of it's single-threaded COM object implementations. The threading model in your rgs file was set to 'Free' as that is the only threading model supported in non DCOM Windows CE platforms."
#endif



// CScriptEngine

class ATL_NO_VTABLE CScriptEngine :
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CScriptEngine, &CLSID_ScriptEngine>,
	public ISupportErrorInfo,
	public IActiveScriptSite,
	public IActiveScriptSiteWindow,
	public IDispatchImpl<IScriptEngine, &IID_IScriptEngine, &LIBID_ScriptEngineLib, /*wMajor =*/ 1, /*wMinor =*/ 0>
{
public:
	CScriptEngine();

DECLARE_REGISTRY_RESOURCEID(IDR_SCRIPTENGINE)

BEGIN_COM_MAP(CScriptEngine)
	COM_INTERFACE_ENTRY(IScriptEngine)
	COM_INTERFACE_ENTRY(IDispatch)
	COM_INTERFACE_ENTRY(ISupportErrorInfo)
	COM_INTERFACE_ENTRY(IActiveScriptSite)
	COM_INTERFACE_ENTRY(IActiveScriptSiteWindow)
END_COM_MAP()

// ISupportsErrorInfo
	STDMETHOD(InterfaceSupportsErrorInfo)(REFIID riid);

	DECLARE_PROTECT_FINAL_CONSTRUCT()

	HRESULT FinalConstruct();

	void FinalRelease();

public:
   // IActiveScriptSite

    STDMETHOD(GetLCID)(LCID *plcid);
    STDMETHOD(GetItemInfo)(LPCOLESTR pstrName, DWORD dwReturnMask, IUnknown **ppiunkItem, ITypeInfo **ppti);
    STDMETHOD(GetDocVersionString)(BSTR *pbstrVersion);
    STDMETHOD(OnScriptTerminate)(const VARIANT *pvarResult, const EXCEPINFO *pexcepinfo);
    STDMETHOD(OnStateChange)(SCRIPTSTATE ssScriptState);
    STDMETHOD(OnScriptError)(IActiveScriptError *pIActiveScriptError);
    STDMETHOD(OnEnterScript)(void);
    STDMETHOD(OnLeaveScript)(void);

    // IActiveScriptSiteWindow

    STDMETHOD(GetWindow)(HWND *phWnd);
    STDMETHOD(EnableModeless)(BOOL fEnable);

	// Misc

	STDMETHOD(ParseScriptText)(
		LPCOLESTR pstrCode,
		LPCOLESTR pstrLanguage,
        LPCOLESTR pstrItemName,
        IUnknown *punkContext,
        LPCOLESTR pstrDelimiter,
        DWORD dwSourceContextCookie,
        ULONG ulStartingLineNumber,
        DWORD dwFlags,
        VARIANT *pvarResult,
        EXCEPINFO *pexcepinfo,
		IActiveScript** ppIActiveScript);
	STDMETHOD(AddGlobal)(IDispatch* item) { return m_Globals->AddItem(item); }

	static CComBSTR VBScript;
	static CComBSTR JScript;

protected:
	HWND m_hWnd;
	CComSafeArray<BSTR> m_Names;
	CComSafeArray<VARIANT> m_Values;
	CComSafeArray<BSTR> m_Contexts;
	CAtlMap<BSTR, DWORD> m_Flags;
	HRESULT m_Error;
	CComBSTR m_ErrorString;
	CComPtr<IActiveScript> m_CurrentScript;
	CGlobals* m_Globals;

public:
	STDMETHOD(SetContext)(BSTR Context, DWORD* dwContext = NULL);
	STDMETHOD(Clear)();
	STDMETHOD(CoFree)();
	STDMETHOD(Evaluate)(BSTR ScriptText, BSTR Language, VARIANT* Result);
	STDMETHOD(Execute)(BSTR ScriptText, BSTR Language);
	STDMETHOD(Import)(BSTR Path, BSTR Name, BSTR Language);
	STDMETHOD(ImportScript)(BSTR ScriptText, BSTR Context, BSTR Name, BSTR Language, DWORD dwNameFlags = SCRIPTTEXT_ISVISIBLE);
	STDMETHOD(LoadScript)(BSTR Path, BSTR* ScriptText);
	STDMETHOD(RunScript)(BSTR Path, BSTR Language);
	STDMETHOD(SetItem)(BSTR Name, VARIANT* Object = NULL, DWORD dwFlags = SCRIPTTEXT_ISVISIBLE, LONG* Index = NULL);
	STDMETHOD(SetWindow)(OLE_HANDLE hWnd);

};

OBJECT_ENTRY_AUTO(__uuidof(ScriptEngine), CScriptEngine)
