<html>

<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�������Ҫ�������������</title>
</head>

<body>

<form method="POST" action="<%=Request.ServerVariables("SCRIPT_NAME")%>">
	<p>�������Ҫ������������ƣ�<input type="text" name="nameText" size="26"></p>
	<p><input type="submit" value="�ύ" name="B1"><input type="reset" value="����" name="B2"></p>
</form>
<%
On Error Resume Next
str=Request.Form("nameText")
str=trim(str)
If len(str)>0 Then
	Set obj=Server.CreateObject(str)
	If IsObject(obj) Then
		Response.write "<H><font  size='5' color='#0000FF'>����"&str&"�Ѿ�������</font>"
	Else
		Response.write "<H><font  size='5' color='#0000FF'>����"&str&"����ʧ�ܣ�</font>"
	End If
End If
%>
</body>

</html>