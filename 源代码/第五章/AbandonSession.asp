<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�½���ҳ 1</title>
</head>
<body>
<%
'����Session�����ֵ
Session("Name")="ASP"
'����Abandon����ɾ��Session�����е�����ֵ
Session.Abandon
'���Session�����ֵ
Response.write "Session('Name')��ֵΪ��"&Session("Name")
%>
<a href="ReDisplay.asp">ReDisplay�ļ�</a>
</body>
</html>
