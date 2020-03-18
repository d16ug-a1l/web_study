<%
Application.Lock
Application("Counter")=Application("Counter")+1
Application.UnLock
%>
<html>

<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>计数器</title>
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
<p align="center"><b><font face="宋体" size="6" color="#0000FF">欢迎你访问本站</font></b></p>
<p align="center"><b><font face="华文行楷" size="6" color="#0000FF">你是本站第<%=str%>位访问者</font></b></p>
</body>
</html>