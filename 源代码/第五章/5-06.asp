<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>本次会话</title>
</head>
<body>
<% 
Session.LCID=2052
'输出当前系统的LCID的值，当前系统为简体中文
Response.write "当前系统的LCID为："&Session.LCID&"<BR>"
'以下输出系统的时间和货币格式
Response.write "时间格式为："&now()&"<BR>"
response.write "货币格式为："&CCur(12360.12)&"<BR>"
Response.write "<BR>"
'设置LCID为繁体中文
Session.LCID=1028
'输出繁体中文的时间和货币的显示格式。
response.write "繁体中文的LCID为："&Session.LCID&"<BR>"
Response.write "时间格式为："&now()&"<BR>"
response.write "货币格式为："&CCur(12360.12)&"<BR>"
%>
</body>
</html>
