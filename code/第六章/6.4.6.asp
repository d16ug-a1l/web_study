<html>
<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>留言内容</title>
</head>
<body>
<form method="POST" action="<%=Request.ServerVariables("SCRIPT_NAME")%>">
	<p>删除文件名称：<input type="text" name="nameText" size="26"></p>
	<p ><input type="submit" value="提交" name="B1"><input type="reset" value="重置" name="B2"></p>
</form>
<%
On Error Resume Next							'启动错误处理程序
'str保存用户输入文件名称，SourcePath保存该文件的物理路径
Dim str,SourcePath
str=Request.Form("nameText")
str=trim(str)
SourcePath=Server.MapPath(str)
'判断用户输入的内容是否为空，不为空则进行删除操作
response.write(SourcePath)
If len(str)>0 Then
	Dim obj
	set obj=Server.CreateObject("Scripting.FileSystemObject")
	'指定文件存在则删除
	
	If obj.FileExists(SourcePath) Then
		obj.DeleteFile(SourcePath)
		Response.write str&"文件已经删除"
		set obj=nothing
	End IF
Else
	Response.write "请输入文件的名称。"
End If
%>
</body>
</html>
