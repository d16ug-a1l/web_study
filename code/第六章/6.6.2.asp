<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>显示错误</title>
</head>
<body>
<%
'On Error Resume Next
str="显示错误信息"
n=cint(str)
If Err.Number>0 Then
	Response.write "发生错误。<BR>"
	Response.write " 错误代号："&Err.Number&"<BR>"
	Response.write "错误原因："&Err.Description&"<BR>"
End If
%>
</body>
</html>
