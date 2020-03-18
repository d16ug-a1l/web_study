<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Page Counter</title>
</head>

<body>
<%
Set Obj=server.CreateObject("MSWC.Counters")
Obj.Increment("my_Count")
n=Cint(Obj.Get("my_Count"))
If n=1 or (n mod 10) =0 Then
	Response.write("<script Language='JavaScript'>alert('恭喜你是我站第"&n&"位访问者！')</script>")
End If
Response.write("访问次数为："& ShowPhoto (n))
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
