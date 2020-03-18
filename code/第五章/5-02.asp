<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>新建网页 1</title>
</head>
<body>
<%
'把创建的对象实例保存在Session对象中
set Session("ObjConn")=Server.CreateObject("scripting.filesystemobject")
'把保存有对象实例的Session对象取出并赋给变量obj
set obj=Session("ObjConn")
'使用obj打开文件txt.txt
set objfile=obj.opentextfile(server.Mappath("txt.txt"),1,true)
'输出文本文件的内容
response.write objfile.readall
%>
</body>
</html>
