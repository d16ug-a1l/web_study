<%
'设置脚本的超时时间为10秒钟。
Server.scriptTimeout=10
%>
<html>
<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>新建网页 1</title>
</head>
<body>
<%
dim start
start=100
'重复100次，以等待超过10秒钟检验脚本超时时间
for k=1 To start
	'获取当前时间的秒数
	nexttime=dateadd("s",1,time)
	'设置停留一秒钟。原理是获取当前的秒数，使用循环等待时间过去一秒钟
	do while time<nexttime
	loop
	'在当前行输出k个“*”
	for i=1 To k
		response.write "*"   '输出信息
	Next
	response.write "<BR>"
next
%>
</body>
</html>
