<html>
<%

If Not(Session("Pass") = True and  Session("User") <>"" and Session("Id") <>"" and Session("GroupID")<>"" )Then
   Response.Redirect("logon.asp")
End If
UserID=Cint(Session("Id"))
GroupID=Cint(Session("GroupID"))
If GroupID<0 Then
	'Response.Redirect("logon.asp") 
	Response.Write("���û�û��Ȩ�����ã������µ�½")
	
End If

%>
<head>
<meta HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=gb2312">
<title>�½���ҳ 1</title>
</head>

<frameset cols="200,*">
	<frame name="contents" target="main" src="left.asp">
	<frame name="main" src="right.asp" target="_self">
	<noframes>
	<body>

	<p>����ҳʹ���˿�ܣ��������������֧�ֿ�ܡ�</p>

	</body>
	</noframes>
</frameset>

</html>