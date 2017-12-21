import "json2.js"
Set jsw = load("json_wrapper.js")


MsgBox "Hello3 " & now

Set obj = jsw.parse("{ ""a"" : [2,3,5,7] } ")

REM obj = jsw.parse("{ ""a"" : [2,3,5,7] } ")

MsgBox "Hello3 " & jsw.stringify(obj)

