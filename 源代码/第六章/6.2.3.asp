<html>
<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>��������</title>
</head>
<body>
<form method="POST" action="<%=Request.ServerVariables("SCRIPT_NAME")%>">
	<p>�������ݣ�<input type="text" name="nameText" size="26"></p>
	<p ><input type="submit" value="�ύ" name="B1"><input type="reset" value="����" name="B2"></p>
</form>
<%
On Error Resume Next
str=Request.Form("nameText")
str=trim(str)
If len(str)>0 Then
	Response.write "�û����������ǣ�"&Server.HTMLEncode(str)
Else
	Response.write "�û�û�����ԡ�"
End If
%>
</body>
</html>