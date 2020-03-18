<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>新建网页 1</title>
</head>

<body>
<form method="POST" action="9.2.1.asp?n=abandon">
	<p><input type="submit" value="提交" name="B1"><input type="reset" value="重置" name="B2"></p>
</form>
<%
Response.write "你在本网站停留了："&Session("nTime")&"分钟。"
n=Trim(Request("n"))
If n="abandon" Then
	Session.Abandon
End If
%>
</body>

</html>