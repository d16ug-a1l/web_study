<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title></title>
</head>
<body>
<%
Dim item
item=Request.Form("ContentRemove")
all=trim(Request.Form("C1"))
Application.Lock
If all="ON" Then 
Response.write all
	Application.Contents.Removeall
Else
	Application.Contents.Remove(item)
End If
Application.unlock
%>
</body>
</html>
