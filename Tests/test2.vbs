
For i = 1 to 5
  Dim se
  Set se = CreateObject("Scripting.ScriptEngine")

  WScript.Echo "Load"

  Dim dom
  Set dom = CreateObject("Microsoft.XMLDOM")
  se.SetItem "dom",dom
  REM Set dom = Nothing

  WScript.Echo "Free"

  se.CoFree
  se.Clear

  WScript.Sleep 5000
Next

