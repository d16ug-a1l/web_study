<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�½���ҳ 1</title>
</head>
<a href="index.asp?id=3&name=admin">link</a>
<body>
<%
'ʹ��ѭ����ȡURL�в�����ֵ
For each str in  Request.QueryString
	Response.write str&"ֵ��"&Request.QueryString(str)&"<BR>"
Next
%>
</body>
</html>
