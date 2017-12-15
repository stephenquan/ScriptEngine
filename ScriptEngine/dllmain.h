// dllmain.h : Declaration of module class.

class CScriptEngineLibModule : public CAtlDllModuleT< CScriptEngineLibModule >
{
public :
	DECLARE_LIBID(LIBID_ScriptEngineLibLib)
	DECLARE_REGISTRY_APPID_RESOURCEID(IDR_SCRIPTINGLIB, "{2A56EE88-E3F3-43F3-8B5B-6320A793EF54}")
};

extern class CScriptEngineLibModule _AtlModule;
