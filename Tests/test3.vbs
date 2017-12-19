
Dim se 
Set se = CreateObject("Scripting.ScriptEngine")

str = "MsgBox Now" & vbCrLf
se.RunScript "test3_extra.vbs", "VBScript"
REM se.Execute str, "VBScript"
se.CoFree
se.RunScript "test3_extra.vbs", "VBScript"
REM se.Execute str, "VBScript"
se.CoFree
se.RunScript "test3_extra.vbs", "VBScript"
REM se.Execute str, "VBScript"
se.CoFree
se.RunScript "test3_extra.vbs", "VBScript"
REM se.Execute str, "VBScript"
se.CoFree

Set se = Nothing
