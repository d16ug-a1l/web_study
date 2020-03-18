<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Page Counter</title>
</head>

<body>
<%
Set Obj=server.CreateObject("MSWC.PageCounter")
Obj.PageHit
If Obj.Hits=1 or Obj.Hits mod 10 =0 Then
	Response.write("<script Language='JavaScript'>alert('恭喜你是我站第"&Obj.Hits&"位访问者！')</script>")
End If
Response.write("访问次数为："& ShowPhoto (Obj.Hits))
Function ShowPhoto(num)
	Dim result
	Str=CStr(num)
	For i=1 to len(str)
		strPhoto=Mid(str,I,1)
		result=result &"<image src='"& strPhoto&".gif' alt="& strPhoto&">"
Next
ShowPhoto=result
End Function
%>
</body>

</html>
