<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>本次会话</title>
</head>
<body>
系统默认会话保留时间为：
<% 
'输出系统当前的会话保留时间，以分钟为单位
Response.write Session.timeout&"分钟<BR>"
'设置系统当前的系统会话保留时间为一小时
Session.TimeOUt=60
Response.write "系统当前会话保留时间为："&Session.timeout&"分钟。"
%>
</body>
</html>
