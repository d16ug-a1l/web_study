<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�½���ҳ 1</title>
</head>
<body>
<%
Response.write Session("Path")&"<BR>"
'����ѭ����ʾ���еĶ���
For Each obj in Session.StaticObjects
	If IsObject(Session.StaticObjects(obj)) Then
		Response.Write "OBJECTԪ��IDΪ�� "&obj &"<br>"
	End If
Next
%>
</body>
</html>
