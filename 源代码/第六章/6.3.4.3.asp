<html>
<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>��������</title>
</head>
<body>
<form method="POST" action="<%=Request.ServerVariables("SCRIPT_NAME")%>">
	<p>ɾ���ļ������ƣ�<input type="text" name="nameText" size="26"></p>
	<p ><input type="submit" value="�ύ" name="B1"><input type="reset" value="����" name="B2"></p>
</form>
<%
On Error Resume Next
Dim str,SourcePath
str=Request.Form("nameText")
str=trim(str)
SourcePath=Server.MapPath(str)
If len(str)>0 Then
	Dim obj
	set obj=Server.CreateObject("Scripting.FileSystemObject")
	If obj.FolderExists(SourcePath) Then
		obj.DeleteFolder(SourcePath)
		Response.write str&"�ļ����Ѿ�ɾ��"
		set obj=nothing
	End IF
Else
	Response.write "�������ļ��е����ơ�"
End If
%>
</body>
</html>