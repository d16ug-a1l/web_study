<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>���λỰ</title>
</head>
<body>
<% 
Session.LCID=2052
'�����ǰϵͳ��LCID��ֵ����ǰϵͳΪ��������
Response.write "��ǰϵͳ��LCIDΪ��"&Session.LCID&"<BR>"
'�������ϵͳ��ʱ��ͻ��Ҹ�ʽ
Response.write "ʱ���ʽΪ��"&now()&"<BR>"
response.write "���Ҹ�ʽΪ��"&CCur(12360.12)&"<BR>"
Response.write "<BR>"
'����LCIDΪ��������
Session.LCID=1028
'����������ĵ�ʱ��ͻ��ҵ���ʾ��ʽ��
response.write "�������ĵ�LCIDΪ��"&Session.LCID&"<BR>"
Response.write "ʱ���ʽΪ��"&now()&"<BR>"
response.write "���Ҹ�ʽΪ��"&CCur(12360.12)&"<BR>"
%>
</body>
</html>
