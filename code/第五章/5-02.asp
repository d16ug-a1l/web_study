<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�½���ҳ 1</title>
</head>
<body>
<%
'�Ѵ����Ķ���ʵ��������Session������
set Session("ObjConn")=Server.CreateObject("scripting.filesystemobject")
'�ѱ����ж���ʵ����Session����ȡ������������obj
set obj=Session("ObjConn")
'ʹ��obj���ļ�txt.txt
set objfile=obj.opentextfile(server.Mappath("txt.txt"),1,true)
'����ı��ļ�������
response.write objfile.readall
%>
</body>
</html>
