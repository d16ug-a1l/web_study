<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�½���ҳ 1</title>
</head>

<body>
<form method="POST" action="9.2.1.asp?n=abandon">
	<p><input type="submit" value="�ύ" name="B1"><input type="reset" value="����" name="B2"></p>
</form>
<%
Response.write "���ڱ���վͣ���ˣ�"&Session("nTime")&"���ӡ�"
n=Trim(Request("n"))
If n="abandon" Then
	Session.Abandon
End If
%>
</body>

</html>