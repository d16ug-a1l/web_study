<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�½���ҳ 1</title>
</head>
<body>
<%
'��������ASPContent����ֵ
Dim ASPContent(4)
ASPContent(0)="ASP��ȫĿ¼"
ASPContent(1)="ASP���ö���"
ASPContent(2)="ASPFSO����"
ASPContent(3)="ASPADO����"
ASPContent(4)="ASP��վ��ȫ"
'��ASPContent�����ֵ����Session������
Session("ASP")=ASPContent
'��Session�����л�ȡ���е�����
strASP=Session("ASP")
'���������е�����Ԫ�ز����
for each str in strASP
	Response.write str&"<BR>"
Next
%>
</body>
</html>
