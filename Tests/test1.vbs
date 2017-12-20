Dim se
Set se = CreateObject("Scripting.ScriptEngine")

REM se.Import "test1_func.vbs", "test1_func.vbs", "VBScript"
REM Import "test1_func.js", "test1_func.js", "JScript"
se.Import "test1_func.vbs"
se.Import "test1_func.js"
WScript.Echo se.Evaluate("XYZZY()", "VBScript")
WScript.Echo se.Evaluate("XYZZY()", "VBScript")
WScript.Echo se.Evaluate("js_date()", "VBScript")
