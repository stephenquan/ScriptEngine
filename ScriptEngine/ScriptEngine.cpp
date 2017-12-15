// ScriptEngine.cpp : Implementation of CScriptEngine

#include "stdafx.h"
#include "ScriptEngine.h"


// CScriptEngine

CScriptEngine::CScriptEngine() :
	m_hWnd(NULL)
{
	m_Names.Create();
	m_Values.Create();

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
	CHECKHR(SetItem(CComBSTR("Engine"), &CComVariant((IDispatch*) this)));
	return S_OK;
}

// IActiveScriptSite

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

	m_ErrorString.Append(m_Names.GetAt(dwSourceContext));
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

	OutputDebugStringW((BSTR) m_ErrorString);
	//wprintf(L"%s\n", (BSTR) m_ErrorString);

	m_Error = ei.scode;

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

	CHECKHR(m_Names.Destroy());
	CHECKHR(m_Values.Destroy());
	CHECKHR(m_Names.Create());
	CHECKHR(m_Values.Create());

	return S_OK;

}

STDMETHODIMP CScriptEngine::ParseScriptText(
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
		DWORD dwFlags = SCRIPTITEM_ISPERSISTENT | SCRIPTITEM_ISVISIBLE | SCRIPTITEM_GLOBALMEMBERS;
		CHECKHR(spIActiveScript->AddNamedItem(name, dwFlags));
	}

	CComPtr<IActiveScriptParse> spIActiveScriptParse;
	CHECKHR(spIActiveScript->QueryInterface(&spIActiveScriptParse));
	CHECKHR(spIActiveScriptParse->InitNew());

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

	LPCOLESTR pstrCode = (LPCOLESTR) ScriptText;
	LPCOLESTR pstrLanguage = (LPCOLESTR) Language;
    LPCOLESTR pstrItemName = (LPCOLESTR) NULL;
    IUnknown *punkContext = NULL;
    LPCOLESTR pstrDelimiter = NULL;
    DWORD dwSourceContextCookie = 0;
    ULONG ulStartingLineNumber = 0;
	DWORD dwFlags = SCRIPTTEXT_ISEXPRESSION;

	EXCEPINFO ei = { };
	CComPtr<IActiveScript> spIActiveScript;
	CHECKHR(ParseScriptText(pstrCode, pstrLanguage, pstrItemName, punkContext, pstrDelimiter, dwSourceContextCookie, ulStartingLineNumber, dwFlags, Result, &ei, &spIActiveScript));

	return hr;
}

//

STDMETHODIMP CScriptEngine::Execute(BSTR ScriptText, BSTR Language)
{
	HRESULT hr = S_OK;

	LPCOLESTR pstrCode = (LPCOLESTR) ScriptText;
	LPCOLESTR pstrLanguage = (LPCOLESTR) Language;
    LPCOLESTR pstrItemName = (LPCOLESTR) NULL;
    IUnknown *punkContext = NULL;
    LPCOLESTR pstrDelimiter = NULL;
    DWORD dwSourceContextCookie = 0;
    ULONG ulStartingLineNumber = 0;
	DWORD dwFlags = 0;
	CComVariant result;

	EXCEPINFO ei = { };
	CComPtr<IActiveScript> spIActiveScript;
	CHECKHR(ParseScriptText(pstrCode, pstrLanguage, pstrItemName, punkContext, pstrDelimiter, dwSourceContextCookie, ulStartingLineNumber, dwFlags, &result, &ei, &spIActiveScript));

	return hr;
}

//

STDMETHODIMP CScriptEngine::Import(BSTR Path, BSTR Name, BSTR Language)
{
	HRESULT hr = S_OK;
	BOOL ok = TRUE;

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

	return ImportScript((BSTR) bytesW, Path, Name, Language);
}

//

STDMETHODIMP CScriptEngine::ImportScript(BSTR scriptText, BSTR Context, BSTR name, BSTR Language)
{
	HRESULT hr = S_OK;

	LPCOLESTR pstrCode = (LPCOLESTR) scriptText;
	LPCOLESTR pstrLanguage = (LPCOLESTR) Language;
    LPCOLESTR pstrItemName = (LPCOLESTR) NULL;
    IUnknown *punkContext = NULL;
    LPCOLESTR pstrDelimiter = NULL;
    DWORD dwSourceContextCookie = 0;
    ULONG ulStartingLineNumber = 0;
    DWORD dwFlags = SCRIPTTEXT_ISPERSISTENT;
	//DWORD dwFlags = SCRIPTTEXT_ISVISIBLE;
	//DWORD dwFlags = 0; // SCRIPTTEXT_ISEXPRESSION;
	CComVariant result;
	EXCEPINFO ei = { };
	CComPtr<IActiveScript> spIActiveScript;

	LONG Index = -1;
	CHECKHR(SetItem(name, NULL, &Index));
	dwSourceContextCookie = Index;

	CHECKHR(m_Names.Add(name));

	CHECKHR(ParseScriptText(pstrCode, pstrLanguage, pstrItemName, punkContext, pstrDelimiter, dwSourceContextCookie, ulStartingLineNumber, dwFlags, &result, &ei, &spIActiveScript));

	CComPtr<IDispatch> spIDispatch;
	CHECKHR(spIActiveScript->GetScriptDispatch(NULL, &spIDispatch));

	CHECKHR(SetItem(name, &CComVariant((IDispatch*) spIDispatch)));

	return hr;
}

STDMETHODIMP CScriptEngine::SetItem(BSTR Name, VARIANT* Object, LONG* Index)
{
	HRESULT hr = S_OK;

	if (Index) *Index = -1;

	for (ULONG idx = 0; idx < m_Names.GetCount() && idx < m_Values.GetCount(); idx++)
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

	return S_OK;
}

STDMETHODIMP CScriptEngine::SetWindow(OLE_HANDLE hWnd)
{
	m_hWnd = (HWND) hWnd;
	return S_OK;
}

