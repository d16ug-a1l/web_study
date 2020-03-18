<!--#include file="adovbs.inc"-->
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>新建网页 1</title>
</head>

<body>
<%
'创建ADO DB.Connection对象
Set Conn=Server.Createobject("Adodb.Connection") 

'依据连接的数据库设置连接字符串
Conn.ConnectionString="Provider=Microsoft.Jet.OLEDB.4.0;"&_
			 			"Data Source="&Server.MapPath("oa.mdb")
Conn.Open  '打开与数据库的连接
Set rs=Server.CreateObject("ADODB.Recordset")
rs.Open "Group_Info",Conn,adOpenKeyset,adLockOptimistic,adCmdTable
'设置每页显示记录的数目
rs.PageSize=2
Dim page
page=rs.PageCount
PageNo=Trim(Request.QueryString("Page"))
If PageNo="" Then PageNo=1
PageNo=Cint(PageNo)
If PageNo<1 Then PageNo=1
If PageNo>Page Then PageNo=Page
rs.AbsolutePage=PageNo
For i=1 To rs.PageSize
	If rs.EOF then exit for
	Response.write "<font color='#FF0000'>职位：</font>"&rs("Name")&_
				"。<font color='#FF0000'>描述信息：</font>"&rs("Info")&"<BR>"
	rs.movenext
next
For i=1 to Page
	Response.write "<a href='13.3.2.asp?Page="&i&"'>第"&i&"页</a>  "
Next
rs.close
Set rs=Nothing
Conn.Close
Set Conn=Nothing
%>
</body>

</html>
