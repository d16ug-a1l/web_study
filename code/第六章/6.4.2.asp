<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�½���ҳ 1</title>
</head>

<body>
<%
Set fs=server.createObject("Scripting.FileSystemObject")
Set file=fs.CreateTextFile("F:\test.htm",true)
file.close
Response.write "�ļ�������ϣ�"
%>
</body>

</html>
