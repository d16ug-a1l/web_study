<!--#include file="funciton.asp"-->
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>权 限 管 理 模 块</title>
<base target="main">
</head>

<body bgcolor="#99CCFF">

<table border="1" width="100%">
<tr><td><font face="华文行楷" size="6" color="#0000FF">栏 目</font></td></tr>
<%
Set Conn=Server.Createobject("Adodb.Connection") 
Conn.ConnectionString="Provider=Microsoft.Jet.OLEDB.4.0;"&_
 			"Data Source="&Server.MapPath("user.mdb")
Conn.Open
UserID=Cint(Session("Id"))
GroupID=Cint(Session("GroupID"))
Sql="Select * From Res_Info "
Set rs1=Conn.Execute(Sql)
nCount=0
do while rs1.EOF=False
	strName= rs1("Name")
	nCode=GetResAllow(strName,UserID,GroupID)
	If nCode>1 Then
		nCount=nCount+1
		Response.write("<tr><td>")
		Response.write("<a href='right.asp?ID="&rs1("ID")&"&Type=LM'>"&rs1("Name")&"</a>")
		Response.write("</td></tr>")
	End If
	rs1.movenext
loop
If nCount=0 Then
	Response.write("<tr><td>")
	Response.write("暂无该栏目的信息！")
	Response.write("</td></tr>")
End If
%>
	
 
</table>
<table border="1" width="100%">
<tr><td><font face="华文行楷" size="6" color="#0000FF">文章</font></td></tr>
<%
 
Sql="Select * From File_Info "
Set rs1=Conn.Execute(Sql)
nCount=0
do while rs1.EOF=False  
	nCode=GetFileAllow(rs1("ID"),UserID,GroupID,rs1("Owner"),rs1("Allow"))
    If nCode>1 Then
    	nCount=nCount+1
		Response.write("<tr><td>")
		Response.write("<a href='right.asp?ID="&rs1("ID")&"&Type=File&Allow="&rs1("Allow")&"&Owner="&rs1("Owner")&"'>"&rs1("Content")&"</a>")
		Response.write("</td></tr>")
	End If
	rs1.movenext
loop
If nCount=0 Then
	Response.write("<tr><td>")
	Response.write("暂无该栏目的信息！")
	Response.write("</td></tr>")
End If
%>
	
 
</table>

</body>

</html>
