<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>新建网页 1</title>
</head>
<a href="index.asp?id=3&name=admin">link</a>
<body>
<%
'使用循环获取URL中参数的值
For each str in  Request.QueryString
	Response.write str&"值是"&Request.QueryString(str)&"<BR>"
Next
%>
</body>
</html>
