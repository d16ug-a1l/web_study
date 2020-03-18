<html>
<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>留言内容</title>
</head>
<body>
<form method="POST" action="<%=Request.ServerVariables("SCRIPT_NAME")%>">
	<p>创建文件夹名称：<input type="text" name="nameText" size="26"></p>
	<p ><input type="submit" value="提交" name="B1"><input type="reset" value="重置" name="B2"></p>
</form>
<%
On Error Resume Next
str=Request.Form("nameText")
str=trim(str)
If len(str)>0 Then
	Dim obj
	set obj=Server.CreateObject("Scripting.FileSystemObject")
	obj.CreateFolder(Server.MapPath("\"&str))
	Response.write str&"文件夹已经建立"
	set obj=nothing
Else
	Response.write "请输入文件夹的名称。"
End If
%>
</body>
</html>