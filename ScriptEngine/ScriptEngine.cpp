// ScriptEngine.cpp : Implementation of CScriptEngine

#include "stdafx.h"
#include "ScriptEngine.h"


// CScriptEngine

CComBSTR CScriptEngine::VBScript("VBScript");
CComBSTR CScriptEngine::JScript("JScript");

CScriptEngine::CScriptEngine() :
	m_hWnd(NULL),
	m_Globals(NULL)
{
	m_Names.Create();
	m_Values.Create();
	m_Contexts.Create();
}

STDMETHODIMP CScriptEngine::InterfaceSupportsErrorInfo(REFIID riid)
{
	static const IID* arr[] = 
	{
		&IID_IScriptEngine
	};

	for (int i=0; i < sizeof(arr) / sizeof(arr[0]); i++)
	{
		if (InlineIsEqualGUID(*arr[i],riid))
			return S_OK;
	}
	return S_FALSE;
}

HRESULT CScriptEngine::FinalConstruct()
{
	HRESULT hr = S_OK;
	m_Globals = new CGlobals();
	//CHECKHR(SetItem(CComBSTR("Engine"), &CComVariant((IDispatch*) this)));
	DWORD dwFlags = SCRIPTITEM_ISVISIBLE | SCRIPTITEM_GLOBALMEMBERS;
	CHECKHR(SetItem(CComBSTR("Globals"), &CComVariant(m_Globals), dwFlags));
	CHECKHR(AddGlobal(&CComVariant(this)));
	return S_OK;
}

void CScriptEngine::FinalRelease()
{
	m_Globals->Release();
	m_Globals = NULL;
}

// IActiveScriptSite

STDMETHODIMP CScriptEngine::SetContext(LPCOLESTR Context, DWORD* dwContext)
{
	HRESULT hr = S_OK;

	if (!Context || !*Context)
	{
		if (dwContext) *dwContext = 0;
		return S_OK;
	}

	for (ULONG idx = 0; idx < m_Contexts.GetCount(); idx++)
	{
		if (_wcsicmp(Context, m_Contexts.GetAt(idx)) == 0)
		{
			if (dwContext)
			{
				*dwContext = idx + 256;
			}
			return S_OK;
		}
	}

	if (dwContext)
	{
		*dwContext = m_Contexts.GetCount() + 256;
	}

	CHECKHR(m_Contexts.Add(CComBSTR(Context)));

	return S_OK;
}

STDMETHODIMP CScriptEngine::GetLCID(LCID *plcid)
{
	*plcid = 0;
	return S_OK;
}

STDMETHODIMP CScriptEngine::GetItemInfo(LPCOLESTR pstrName, DWORD dwReturnMask, IUnknown **ppiunkItem, ITypeInfo **ppti)
{
	HRESULT hr = S_OK;

	for (ULONG idx = 0; idx < m_Names.GetCount() && idx < m_Values.GetCount(); idx++)
	{
		BSTR name = m_Names.GetAt(idx);

		if (wcscmp(name, pstrName) == 0)
		{
			if (dwReturnMask & SCRIPTINFO_IUNKNOWN)
			{
				*ppiunkItem = NULL;
				VARIANT& Value = m_Values.GetAt(idx);
				if (Value.vt != VT_DISPATCH) return TYPE_E_ELEMENTNOTFOUND;
				CHECKHR(V_DISPATCH(&Value)->QueryInterface(IID_IUnknown, (void**) ppiunkItem));
			}

			return S_OK;
		}
	}

	return TYPE_E_ELEMENTNOTFOUND;
}

STDMETHODIMP CScriptEngine::GetDocVersionString(BSTR *pbstrVersion)
{
	*pbstrVersion = SysAllocString(L"1.0");
	return S_OK;
}

STDMETHODIMP CScriptEngine::OnScriptTerminate(const VARIANT *pvarResult, const EXCEPINFO *pexcepinfo)
{
	return S_OK;
}

STDMETHODIMP CScriptEngine::OnStateChange(SCRIPTSTATE ssScriptState)
{
	return S_OK;
}

STDMETHODIMP CScriptEngine::OnScriptError(IActiveScriptError *pIActiveScriptError)
{
	HRESULT hr = S_OK;

	CComBSTR m_ErrorString;
	WCHAR szText[1024] = { };

	EXCEPINFO ei = { };
	hr = pIActiveScriptError->GetExceptionInfo(&ei);

	DWORD dwSourceContext = 0;
	ULONG ulLineNumber = 0;
	LONG lCharacterPosition = 0;
	hr = pIActiveScriptError->GetSourcePosition(&dwSourceContext, &ulLineNumber, &lCharacterPosition);

	BSTR name = (dwSourceContext >= 256) ? m_Contexts.GetAt(dwSourceContext - 256) : NULL;
	m_ErrorString.Append(name);
	::wsprintf(szText, L"(%d, %d)", ulLineNumber + 1, lCharacterPosition + 1);
	m_ErrorString.Append(szText);

	if (ei.bstrSource)
	{
		m_ErrorString.Append(L" ");
		m_ErrorString.Append(ei.bstrSource);
	}

	if (ei.bstrDescription)
	{
		m_ErrorString.Append(L": ");
		m_ErrorString.Append(ei.bstrDescription);
	}

	CComBSTR bstrSourceLine;
	CHECKHR(pIActiveScriptError->GetSourceLineText(&bstrSourceLine));
	if ((BSTR) bstrSourceLine)
	{
		m_ErrorString.Append(L": ");
		m_ErrorString.Append((BSTR) bstrSourceLine);
	}

	m_Error = ei.scode;

	OutputDebugStringW((BSTR) m_ErrorString);
	OutputDebugStringW(L"\r\n");
	//wprintf(L"%s\n", (BSTR) m_ErrorString);

	CComPtr<ICreateErrorInfo> spICreateErrorInfo;
	CHECKHR(CreateErrorInfo(&spICreateErrorInfo));
	spICreateErrorInfo->SetDescription((BSTR) m_ErrorString);
	CComPtr<IErrorInfo> spIErrorInfo;
	spICreateErrorInfo->QueryInterface(IID_IErrorInfo, (void**) &spIErrorInfo);
	SetErrorInfo(0, spIErrorInfo);

	if (m_hWnd)
	{
		::MessageBox(m_hWnd, (BSTR) m_ErrorString, L"Error", MB_ICONERROR | MB_OK);
	}

	return S_OK;
}

STDMETHODIMP CScriptEngine::OnEnterScript(void)
{
	return S_OK;
}

STDMETHODIMP CScriptEngine::OnLeaveScript(void)
{
	return S_OK;
}

// IActiveScriptSiteWindow

STDMETHODIMP CScriptEngine::GetWindow(HWND *phWnd)
{
	*phWnd = m_hWnd;
	return S_OK;
}

STDMETHODIMP CScriptEngine::EnableModeless(BOOL fEnable)
{
	return S_OK;
}

// Misc

//

STDMETHODIMP CScriptEngine::Clear()
{
	HRESULT hr = S_OK;

	for (ULONG idx = 0; idx < m_Names.GetCount(); idx++)
	{
		VARIANT& value = m_Values.GetAt(idx);

		if (value.vt == VT_DISPATCH)
		{
			CComPtr<IActiveScript> spIActiveScript;
			hr = V_DISPATCH(&value)->QueryInterface(IID_IActiveScript, (void**) &spIActiveScript);
			if (SUCCEEDED(hr))
			{
				//CHECKHR(spIActiveScript->SetScriptState(SCRIPTSTATE_CLOSED));
				CHECKHR(spIActiveScript->Close());
				spIActiveScript = NULL;
			}
		}

		CHECKHR(VariantClear(&value));
	}

	CHECKHR(m_Names.Destroy());
	CHECKHR(m_Values.Destroy());
	CHECKHR(m_Contexts.Destroy());
	CHECKHR(m_Names.Create());
	CHECKHR(m_Values.Create());
	CHECKHR(m_Contexts.Create());
	
	CHECKHR(m_Globals->Clear());

	::CoFreeUnusedLibrariesEx(0, 0);
	::CoFreeUnusedLibrariesEx(0, 0);

	return S_OK;
}

STDMETHODIMP CScriptEngine::CoFree()
{
	::CoFreeUnusedLibrariesEx(0, NULL);
	::CoFreeUnusedLibrariesEx(0, NULL);
	return S_OK;
}

//

STDMETHODIMP CScriptEngine::ParseScriptText(
	LPCOLESTR pstrCode,
	LPCOLESTR pstrLanguage,
    LPCOLESTR pstrContext,
    DWORD dwFlags,
    VARIANT *pvarResult,
    EXCEPINFO *pexcepinfo,
	IActiveScript** ppIActiveScript)
{
	HRESULT hr = S_OK;

	if (pvarResult)
	{
		VariantInit(pvarResult);
	}

	if (pexcepinfo)
	{
		ZeroMemory(pexcepinfo, sizeof(EXCEPINFO));
	}

	if (ppIActiveScript)
	{
		*ppIActiveScript = NULL;
	}

	CComPtr<IActiveScript> spIActiveScript;
	CHECKHR(spIActiveScript.CoCreateInstance(pstrLanguage));
	CHECKHR(spIActiveScript->SetScriptSite(this));

	for (ULONG idx = 0; idx < m_Names.GetCount(); idx++)
	{
		BSTR name = m_Names.GetAt(idx);
		//DWORD dwFlags = SCRIPTITEM_ISVISIBLE | SCRIPTITEM_GLOBALMEMBERS;
		DWORD dwNameFlags = m_Flags[name];
		CHECKHR(spIActiveScript->AddNamedItem(name, dwNameFlags));
	}

	m_CurrentScript = spIActiveScript;

	CComPtr<IActiveScriptParse> spIActiveScriptParse;
	CHECKHR(spIActiveScript->QueryInterface(&spIActiveScriptParse));
	CHECKHR(spIActiveScriptParse->InitNew());

    LPCOLESTR pstrItemName = NULL;
    DWORD dwSourceContextCookie = 0;
    ULONG ulStartingLineNumber = 0;
    IUnknown *punkContext = NULL;
    LPCOLESTR pstrDelimiter = NULL;

	CHECKHR(SetContext(pstrContext, &dwSourceContextCookie));

	CComVariant result;
	EXCEPINFO ei = { };
	CHECKHR(spIActiveScriptParse->ParseScriptText(pstrCode, pstrItemName, punkContext, pstrDelimiter, dwSourceContextCookie, ulStartingLineNumber, dwFlags, &result, &ei));

	if (pexcepinfo)
	{
		memcpy(pexcepinfo, &ei, sizeof(EXCEPINFO));
	}

	if (pvarResult)
	{
		CHECKHR(result.Detach(pvarResult));
	}

	if (ppIActiveScript)
	{
		*ppIActiveScript = spIActiveScript.Detach();
	}

	m_CurrentScript = NULL;

	return S_OK;
}

//

STDMETHODIMP CScriptEngine::Evaluate(BSTR ScriptText, BSTR Language, VARIANT* Result)
{
	HRESULT hr = S_OK;

	if (Result)
	{
		VariantInit(Result);
	}

	CComBSTR bstrContext("SCRIPT");
	DWORD dwFlags = SCRIPTTEXT_ISEXPRESSION;
	EXCEPINFO ei = { };
	CComPtr<IActiveScript> spIActiveScript;
	CHECKHR(ParseScriptText(ScriptText, Language, (BSTR) bstrContext, dwFlags, Result, &ei, &spIActiveScript));

	CHECKHR(spIActiveScript->Close());

	spIActiveScript = NULL;

	::CoFreeUnusedLibraries();
	::CoFreeUnusedLibraries();

	return hr;
}

//

STDMETHODIMP CScriptEngine::Execute(BSTR ScriptText, BSTR Language)
{
	HRESULT hr = S_OK;

	CComBSTR bstrContext("SCRIPT");
	CComVariant result;
	DWORD dwFlags = 0;
	EXCEPINFO ei = { };
	CComPtr<IActiveScript> spIActiveScript;
	CHECKHR(ParseScriptText(ScriptText, Language, (BSTR) bstrContext, dwFlags, &result, &ei, &spIActiveScript));

	CHECKHR(spIActiveScript->Close());

	::CoFreeUnusedLibrariesEx(0, NULL);
	::CoFreeUnusedLibrariesEx(0, NULL);

	return hr;
}

//

STDMETHODIMP CScriptEngine::Import(BSTR Path, BSTR Language)
{
	HRESULT hr = S_OK;
	CComBSTR ScriptText;
	CHECKHR(LoadTextFile(Path, &ScriptText));
	CHECKHR(ImportScript((BSTR) ScriptText, Path, GetLanguage(Path, Language)));
	return hr;
}

//

STDMETHODIMP CScriptEngine::ImportScript(BSTR scriptText, BSTR Context, BSTR Language)
{
	HRESULT hr = S_OK;

    DWORD dwFlags = SCRIPTTEXT_ISVISIBLE;
	CComVariant result;
	EXCEPINFO ei = { };
	CComPtr<IActiveScript> spIActiveScript;
	CHECKHR(ParseScriptText(scriptText, Language, Context, dwFlags, &result, &ei, &spIActiveScript));
	CHECKHR(AddGlobal(&CComVariant(spIActiveScript)));
	return hr;
}

STDMETHODIMP CScriptEngine::Load(BSTR Path, BSTR Language, IDispatch** Object)
{
	HRESULT hr = S_OK;
	CComBSTR ScriptText;
	CHECKHR(LoadTextFile(Path, &ScriptText));
	CHECKHR(LoadScriptText((BSTR)ScriptText, Path, GetLanguage(Path, Language), Object));
	return hr;
}

STDMETHODIMP CScriptEngine::LoadScriptText(BSTR scriptText, BSTR Context, BSTR Language, IDispatch** Object)
{
	HRESULT hr = S_OK;

	DWORD dwFlags = SCRIPTTEXT_ISVISIBLE;
	CComVariant result;
	EXCEPINFO ei = {};
	CComPtr<IActiveScript> spIActiveScript;
	CHECKHR(ParseScriptText(scriptText, Language, Context, dwFlags, &result, &ei, &spIActiveScript));
	CHECKHR(spIActiveScript->GetScriptDispatch(NULL, Object));
	return hr;
}

STDMETHODIMP CScriptEngine::LoadTextFile(BSTR Path, BSTR* Text)
{
	HRESULT hr = S_OK;
	BOOL ok = TRUE;

	if (!Text) return E_INVALIDARG;
	*Text = NULL;

	LPCWSTR lpFileName = Path;
	DWORD dwDesiredAccess = GENERIC_READ;
    DWORD dwShareMode = FILE_SHARE_READ | FILE_SHARE_WRITE;
    LPSECURITY_ATTRIBUTES lpSecurityAttributes = NULL;
    DWORD dwCreationDisposition = OPEN_EXISTING;
    DWORD dwFlagsAndAttributes = FILE_ATTRIBUTE_NORMAL;
    HANDLE hTemplateFile = NULL;

	HANDLE hFile = CreateFile(lpFileName, dwDesiredAccess, dwShareMode, lpSecurityAttributes, dwCreationDisposition, dwFlagsAndAttributes, hTemplateFile);
	if (hFile == INVALID_HANDLE_VALUE)
	{
		return E_POINTER;
	}

	ULARGE_INTEGER uliFileSize = { };
	uliFileSize.LowPart = GetFileSize(hFile, &uliFileSize.HighPart);

	CComSafeArray<BYTE> bytesA;
	int sizeA = uliFileSize.LowPart + 1;
	hr = bytesA.Create(uliFileSize.LowPart + 1);
	if (FAILED(hr))
	{
		ok = CloseHandle(hFile);
		hFile = INVALID_HANDLE_VALUE;
		return hr;
	}

	ZeroMemory(&bytesA[0], sizeA);

	DWORD dwRead = 0;
	ok = ReadFile(hFile, &bytesA[0], sizeA, &dwRead, NULL);
	if (ok == FALSE)
	{
		hr = HRESULT_FROM_WIN32(GetLastError());
		ok = CloseHandle(hFile);
		hFile = INVALID_HANDLE_VALUE;
		return hr;
	}

	BYTE* pBytesA = &bytesA[0];
	int nLenA = sizeA;

	UINT CodePage = CP_UTF8;
	int nLenW = ::MultiByteToWideChar(CodePage, 0, (LPCSTR) pBytesA, nLenA, NULL, 0);
	if (nLenW < 1)
	{
		return E_OUTOFMEMORY;
	}

	CComBSTR bytesW;
	hr = bytesW.AssignBSTR(SysAllocStringLen(NULL, nLenW - 1));
	if (FAILED(hr) || (BSTR) bytesW == NULL)
	{
		return E_OUTOFMEMORY;
	}

	::MultiByteToWideChar(CodePage, 0, (LPCSTR) pBytesA, nLenA, (BSTR) bytesW, nLenW);

	*Text = bytesW.Detach();

	return S_OK;
}

STDMETHODIMP CScriptEngine::RunScript(BSTR Path, BSTR Language)
{
	HRESULT hr = S_OK;
	BOOL ok = TRUE;

	CComBSTR ScriptText;
	CHECKHR(LoadTextFile(Path, &ScriptText));

	DWORD dwFlags = SCRIPTTEXT_ISVISIBLE;
	CComVariant result;
	EXCEPINFO ei = { };
	CComPtr<IActiveScript> spIActiveScript;
	CHECKHR(ParseScriptText((BSTR) ScriptText, GetLanguage(Path, Language), Path, dwFlags, &result, &ei, &spIActiveScript));

	if (spIActiveScript)
	{
		CHECKHR(spIActiveScript->Close());
		spIActiveScript = NULL;
	}

	//CHECKHR(Clear());

	::CoFreeUnusedLibrariesEx(0, NULL);
	::CoFreeUnusedLibrariesEx(0, NULL);

	return S_OK;
}

STDMETHODIMP CScriptEngine::SetItem(BSTR Name, VARIANT* Object, DWORD dwFlags, LONG* Index)
{
	HRESULT hr = S_OK;

	if (Index) *Index = -1;

	for (ULONG idx = 0; idx < m_Names.GetCount(); idx++)
	{
		if (wcscmp(m_Names.GetAt(idx), Name) == 0)
		{
			if (Index) *Index = (LONG) idx;
			m_Values.SetAt(idx, Object ? *Object : CComVariant());
			return S_OK;
		}
	}

	if (Index) *Index = m_Names.GetCount();
	m_Names.Add(Name);
	m_Values.Add(Object ? *Object : CComVariant());
	m_Flags.SetAt(Name, dwFlags);

	if (m_CurrentScript)
	{
		CHECKHR(m_CurrentScript->AddNamedItem(Name, dwFlags));
	}

	return S_OK;
}

STDMETHODIMP CScriptEngine::SetWindow(OLE_HANDLE hWnd)
{
	m_hWnd = (HWND) hWnd;
	return S_OK;
}

BSTR CScriptEngine::GetLanguage(BSTR Path, BSTR Language)
{
	if (!Language || !*Language)
	{
		if (wcsstr(Path, L".js"))
		{
			Language = (BSTR) JScript;
		}
		else
		{
			Language = (BSTR) VBScript;
		}
	}
	return Language;
}


