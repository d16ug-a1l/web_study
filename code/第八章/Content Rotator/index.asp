<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Content Rotator</title>
</head>

<body>
<%
Set Obj=server.CreateObject("MSWC.ContentRotator")
Response.Write(Obj.ChooseContent("content rotator.txt"))
%>
</body>

</html>
