<html>
<%

If Not(Session("Pass") = True and  Session("User") <>"" and Session("Id") <>"" and Session("GroupID")<>"" )Then
   Response.Redirect("logon.asp")
End If
UserID=Cint(Session("Id"))
GroupID=Cint(Session("GroupID"))
If GroupID<0 Then
	'Response.Redirect("logon.asp") 
	Response.Write("该用户没有权限设置！请重新登陆")
	
End If

%>
<head>
<meta HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=gb2312">
<title>新建网页 1</title>
</head>

<frameset cols="200,*">
	<frame name="contents" target="main" src="left.asp">
	<frame name="main" src="right.asp" target="_self">
	<noframes>
	<body>

	<p>此网页使用了框架，但您的浏览器不支持框架。</p>

	</body>
	</noframes>
</frameset>

</html>