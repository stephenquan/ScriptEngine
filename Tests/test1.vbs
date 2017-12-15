Dim se
Set se = CreateObject("Scripting.ScriptEngine")

se.Import "test1_func.vbs", "test1_func.vbs", "VBScript"
se.Import "test1_func.js", "test1_func.js", "JScript"
WScript.Echo se.Evaluate("XYZZY()", "VBScript")
WScript.Echo se.Evaluate("XYZZY()", "VBScript")
WScript.Echo se.Evaluate("js_date()", "VBScript")
