import "json2.js"
REM import "json_wrapper.js", "sq"
import "json_wrapper.js"


MsgBox "Hello3 " & now

Set obj = json_parse("{ ""a"" : [2,3,5,7] } ")

REM obj = sq.json_parse("{ ""a"" : [2,3,5,7] } ")

MsgBox "Hello3 " & json_stringify(obj)

