<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>新建网页 1</title>
</head>

<body>
<%
Dim Obj
Set Obj=Server.CreateObject("MSWC.AdRotator")
 
Response.Write(Obj.GetAdvertisement("advertise.txt"))
Set Obj=Nothing
%>
</body>

</html>
