<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>新建网页 1</title>
</head>
<body>
<%
Response.write Session("Path")&"<BR>"
'下面循环显示所有的对象
For Each obj in Session.StaticObjects
	If IsObject(Session.StaticObjects(obj)) Then
		Response.Write "OBJECT元素ID为： "&obj &"<br>"
	End If
Next
%>
</body>
</html>
