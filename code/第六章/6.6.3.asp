<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>显示错误</title>
</head>
<body>
<%

Dim objErr
set objErr=Server.GetLastError()
Response.write "错误码："&objErr.ASPCode&"<BR>"
Response.write "错误代号："&objErr.Number&"<BR>"
Response.write "发生错误的程序代码："&objErr.Source&"<BR>"
Response.write "错误行号："&objErr.Line&"<BR>"
Response.write "发生错误的文件："&objErr.File&"<BR>"
Response.write "发生错误的字符："&objErr.Column&"<BR>"
Response.write "错误类型："&objErr.Category&"<BR>"
Response.write "错误描述："&objErr.Description&"<BR>"
Response.write "其他的错误描述："&objErr.ASPDescription&"<BR>"
%>
</body>
</html>

