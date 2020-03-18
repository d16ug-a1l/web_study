<html>
<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>留言内容</title>
</head>
<body>
<form method="POST" action="<%=Request.ServerVariables("SCRIPT_NAME")%>">
	<p>源文件名称：&nbsp; <input type="text" name="SText" size="26"></p>
	<p>目的文件名称：<input type="text" name="DText" size="26"></p>
	<p ><input type="submit" value="提交" name="B1"><input type="reset" value="重置" name="B2"></p>
</form>
<% 
Sstr=Request.Form("SText")
Sstr=trim(Sstr)
Dstr=Request.Form("DText")
Dstr=trim(Dstr)
If len(Sstr)>0 and len(Dstr)>0  Then
	Dim obj
	Dim SourcePath,DestinPath
	DestinPath=Server.MapPath("\"&Dstr)
	SourcePath=Server.MapPath("\"&Sstr)
	set obj=Server.CreateObject("Scripting.FileSystemObject")
	If obj.FileExists(SourcePath) Then
	   obj.CopyFile SourcePath,DestinPath
	   Response.write "文件已经复制完毕"
	End If
	set obj=nothing
Else
	If len(Sstr)>0 Then
		Response.write "请输入源文件的名称。"
	Else 
		Response.write "请输入目的文件的名称。"
	End If
	Response.end
End If
%>
</body>
</html>