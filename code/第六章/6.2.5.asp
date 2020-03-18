<html>
<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>留言内容</title>
</head>
<body>
<%
Response.write "当前文件为："&Request.ServerVariables("SCRIPT_NAME")&"<BR>"
Response.write "该文件的路径为："&Server.MapPath(Request.ServerVariables("SCRIPT_NAME"))&"<BR>"
Response.write "该文件的当前路径为："&Server.MapPath("./")&"<BR>"
Response.write "该文件的父目录路径为："&Server.MapPath("../")&"<BR>"
Response.write "该文件的根目录路径为："&Server.MapPath("/")&"<BR>"
%>
</body>
</html>
