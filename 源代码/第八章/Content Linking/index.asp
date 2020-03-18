<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Content Linking组件</title>
</head>
<body>
<%
Dim Obj
'创建Content Linking组件
Set Obj=Server.CreateObject("MSWC.NextLink")
'获取NextLink.txt文件中的URL数目
Count=Obj.GetListCount("NextLink.txt")
Dim I
'循环显示每个项目
For i=1 to Count 

%>
<p><a href="<%=Obj.GetNthURL("NextLink.txt",i) %>"><%=Obj.GetNthDescription("NextLink.txt",i)%></a></p>
<%
Next
'设置Obj为空
Set Obj=Nothing
%>
</body>
</html>
