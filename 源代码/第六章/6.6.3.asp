<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>��ʾ����</title>
</head>
<body>
<%

Dim objErr
set objErr=Server.GetLastError()
Response.write "�����룺"&objErr.ASPCode&"<BR>"
Response.write "������ţ�"&objErr.Number&"<BR>"
Response.write "��������ĳ�����룺"&objErr.Source&"<BR>"
Response.write "�����кţ�"&objErr.Line&"<BR>"
Response.write "����������ļ���"&objErr.File&"<BR>"
Response.write "����������ַ���"&objErr.Column&"<BR>"
Response.write "�������ͣ�"&objErr.Category&"<BR>"
Response.write "����������"&objErr.Description&"<BR>"
Response.write "�����Ĵ���������"&objErr.ASPDescription&"<BR>"
%>
</body>
</html>

