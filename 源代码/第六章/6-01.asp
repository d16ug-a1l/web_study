<%
'���ýű��ĳ�ʱʱ��Ϊ10���ӡ�
Server.scriptTimeout=10
%>
<html>
<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�½���ҳ 1</title>
</head>
<body>
<%
dim start
start=100
'�ظ�100�Σ��Եȴ�����10���Ӽ���ű���ʱʱ��
for k=1 To start
	'��ȡ��ǰʱ�������
	nexttime=dateadd("s",1,time)
	'����ͣ��һ���ӡ�ԭ���ǻ�ȡ��ǰ��������ʹ��ѭ���ȴ�ʱ���ȥһ����
	do while time<nexttime
	loop
	'�ڵ�ǰ�����k����*��
	for i=1 To k
		response.write "*"   '�����Ϣ
	Next
	response.write "<BR>"
next
%>
</body>
</html>
