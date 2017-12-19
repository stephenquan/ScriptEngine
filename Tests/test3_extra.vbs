Dim dom
Set dom = CreateObject("Microsoft.XMLDOM")
dom.loadXML "<Hello><World/></Hello>"
MsgBox doc.xml
Set dom = Nothing

Dim ado
REM Set ado = CreateObject("ADODB.Connection")
Set ado = Nothing



