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
	Response.write("<script Language='JavaScript'>alert('��ϲ������վ��"&n&"λ�����ߣ�')</script>")
End If
Response.write("���ʴ���Ϊ��"& ShowPhoto (n))
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
