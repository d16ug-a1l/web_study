<%
Application.Lock
Application("Counter")=Application("Counter")+1
Application.UnLock
%>
<html>

<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>������</title>
</head>

<body>
<%
response.write Application("CounterFile")
Dim count
Dim str
str=""
count=Application("Counter")
n=count\10
m=count mod 10
Do while not (n=0 and m=0)
	str="<img src='"&m&".gif' >"&str
	m=n mod 10
	n=n\10	
Loop
%>
<p align="center"><b><font face="����" size="6" color="#0000FF">��ӭ����ʱ�վ</font></b></p>
<p align="center"><b><font face="�����п�" size="6" color="#0000FF">���Ǳ�վ��<%=str%>λ������</font></b></p>
</body>
</html>