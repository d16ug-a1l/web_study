<html>

<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�½���ҳ 1</title>
</head>

<body>
<%Set fs=server.createObject("Scripting.FileSystemObject")
FilePath=Request("FilePath")
Path="D:\�������"
Path=Path&FilePath

if fs.FolderExists(Path) Then
	Set objfolder=fs.GetFolder(Path)
	Set folders=objfolder.subfolders
	response.write "��ǰĿ¼Ϊ<font color=#FF9900>"&objfolder.Name&"</font><BR>"
	If objfolder.IsRootFolder Then
		response.write "��ǰĿ¼Ϊ<font color=#FF9900>��Ŀ¼</font><BR>��"
	Else
		response.write "��ǰĿ¼���Ǹ�Ŀ¼����Ŀ¼Ϊ<font color=#FF9900><BR>"
		response.write objfolder.ParentFolder&"</font>"
	End If
	
	response.write "��ǰĿ¼����Ŀ¼�������ļ���С�ܺ�Ϊ<font color=#FF9900>"&objfolder.Size&"�ֽ�</font><BR>"
	For Each folder In folders
		Response.write "<a href='8.3.3.asp?FilePath="&FilePath&"\"&folder.name&"'>"&_
						folder.name&"</A>"
		Response.write "   "&folder.datecreated&"<BR>"
	Next
	Response.write "<font color=#FF0000>��ǰĿ¼�������ļ�Ϊ��</font><BR>"
	Set Files=objfolder.Files
	For each File in Files
		Response.write "<font color=#FF9900>"&File.name&"</font><BR>"
	Next	
End if

%>
</body>

</html>