<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>��ʾ����</title>
</head>
<body>
<%
'On Error Resume Next
str="��ʾ������Ϣ"
n=cint(str)
If Err.Number>0 Then
	Response.write "��������<BR>"
	Response.write " ������ţ�"&Err.Number&"<BR>"
	Response.write "����ԭ��"&Err.Description&"<BR>"
End If
%>
</body>
</html>
