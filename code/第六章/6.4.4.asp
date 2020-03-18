<html>

<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>故事接龙</title>
</head>

<body>
<%
Content=trim(Request.Form("Content"))
Set fs=server.createObject("Scripting.FileSystemObject")
FileName="test.txt"
FilePath=Server.MapPath(FileName)
If Content<>"" Then
	If not fs.FileExists(FilePath) Then
		Set file=fs.CreateTextFile(FilePath,true)
	Else
		Set file=fs.opentextfile (FilePath,2)
	End If 
	file.writeLine Content&"<BR>"
	file.close
End If
%>
<p align="center">故事接龙</p>
<p align="left" style="margin-top: 0; margin-bottom: 0">故事内容：</p>
<%
If fs.FileExists(FilePath) Then
	set objfile=fs.opentextfile(FilePath,1 )
	text=objfile.Readall
	Response.write text
End If
%>
<p align="left" style="margin-top: 0; margin-bottom: 0">　</p>
<p align="left" style="margin-top: 0; margin-bottom: 0">续借内容为：</p>
<form method="POST" action="8.4.4.asp">
	<p align="center" style="margin-top: 0; margin-bottom: 0">
	<textarea rows="5" name="Content" cols="39"></textarea></p>
	<p align="center" style="margin-top: 0; margin-bottom: 0"><input type="submit" value="提交" name="B1"><input type="reset" value="重置" name="B2"></p>
</form>
<p align="left" style="margin-top: 0; margin-bottom: 0">　</p>

</body>

</html>