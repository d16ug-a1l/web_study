<html>
<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>��������</title>
</head>
<body>
<%
Response.write "��ǰ�ļ�Ϊ��"&Request.ServerVariables("SCRIPT_NAME")&"<BR>"
Response.write "���ļ���·��Ϊ��"&Server.MapPath(Request.ServerVariables("SCRIPT_NAME"))&"<BR>"
Response.write "���ļ��ĵ�ǰ·��Ϊ��"&Server.MapPath("./")&"<BR>"
Response.write "���ļ��ĸ�Ŀ¼·��Ϊ��"&Server.MapPath("../")&"<BR>"
Response.write "���ļ��ĸ�Ŀ¼·��Ϊ��"&Server.MapPath("/")&"<BR>"
%>
</body>
</html>
