<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Content Linking���</title>
</head>
<body>
<%
Dim Obj
'����Content Linking���
Set Obj=Server.CreateObject("MSWC.NextLink")
'��ȡNextLink.txt�ļ��е�URL��Ŀ
Count=Obj.GetListCount("NextLink.txt")
Dim I
'ѭ����ʾÿ����Ŀ
For i=1 to Count 

%>
<p><a href="<%=Obj.GetNthURL("NextLink.txt",i) %>"><%=Obj.GetNthDescription("NextLink.txt",i)%></a></p>
<%
Next
'����ObjΪ��
Set Obj=Nothing
%>
</body>
</html>
