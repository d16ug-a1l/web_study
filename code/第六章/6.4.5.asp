<html>
<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>��������</title>
</head>
<body>
<form method="POST" action="<%=Request.ServerVariables("SCRIPT_NAME")%>">
	<p>Դ�ļ����ƣ�&nbsp; <input type="text" name="SText" size="26"></p>
	<p>Ŀ���ļ����ƣ�<input type="text" name="DText" size="26"></p>
	<p ><input type="submit" value="�ύ" name="B1"><input type="reset" value="����" name="B2"></p>
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
	   Response.write "�ļ��Ѿ��������"
	End If
	set obj=nothing
Else
	If len(Sstr)>0 Then
		Response.write "������Դ�ļ������ơ�"
	Else 
		Response.write "������Ŀ���ļ������ơ�"
	End If
	Response.end
End If
%>
</body>
</html>