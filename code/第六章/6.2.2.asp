<html>

<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>请键入需要创建对象的名称</title>
</head>

<body>

<form method="POST" action="<%=Request.ServerVariables("SCRIPT_NAME")%>">
	<p>请键入需要创建对象的名称：<input type="text" name="nameText" size="26"></p>
	<p><input type="submit" value="提交" name="B1"><input type="reset" value="重置" name="B2"></p>
</form>
<%
On Error Resume Next
str=Request.Form("nameText")
str=trim(str)
If len(str)>0 Then
	Set obj=Server.CreateObject(str)
	If IsObject(obj) Then
		Response.write "<H><font  size='5' color='#0000FF'>对象"&str&"已经创建！</font>"
	Else
		Response.write "<H><font  size='5' color='#0000FF'>对象"&str&"创建失败！</font>"
	End If
End If
%>
</body>

</html>