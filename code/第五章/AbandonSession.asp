<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>新建网页 1</title>
</head>
<body>
<%
'设置Session对象的值
Session("Name")="ASP"
'调用Abandon方法删除Session对象中的所有值
Session.Abandon
'输出Session对象的值
Response.write "Session('Name')的值为："&Session("Name")
%>
<a href="ReDisplay.asp">ReDisplay文件</a>
</body>
</html>
