<html>

<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>新建网页 1</title>
</head>

<body>
<%Set fs=server.createObject("Scripting.FileSystemObject")
FilePath=Request("FilePath")
Path="D:\书稿资料"
Path=Path&FilePath

if fs.FolderExists(Path) Then
	Set objfolder=fs.GetFolder(Path)
	Set folders=objfolder.subfolders
	response.write "当前目录为<font color=#FF9900>"&objfolder.Name&"</font><BR>"
	If objfolder.IsRootFolder Then
		response.write "当前目录为<font color=#FF9900>根目录</font><BR>。"
	Else
		response.write "当前目录不是根目录。父目录为<font color=#FF9900><BR>"
		response.write objfolder.ParentFolder&"</font>"
	End If
	
	response.write "当前目录及子目录中所有文件大小总和为<font color=#FF9900>"&objfolder.Size&"字节</font><BR>"
	For Each folder In folders
		Response.write "<a href='8.3.3.asp?FilePath="&FilePath&"\"&folder.name&"'>"&_
						folder.name&"</A>"
		Response.write "   "&folder.datecreated&"<BR>"
	Next
	Response.write "<font color=#FF0000>当前目录下所有文件为：</font><BR>"
	Set Files=objfolder.Files
	For each File in Files
		Response.write "<font color=#FF9900>"&File.name&"</font><BR>"
	Next	
End if

%>
</body>

</html>