<!--#include file="adovbs.inc"-->
<html>

<head>
<meta http-equiv="Content-Language" content="zh-cn">
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
If not rs.EOF then
	Do while not rs.Eof
		Response.write "<font color='#FF0000'>职位：</font>"&rs("Name")&_
				"。<font color='#FF0000'>描述信息：</font>"&rs("Info")&"<BR>"
		rs.MoveNext
	Loop
Else
	Response.write "没有记录！"
End If
rs.close
Set rs=Nothing
Conn.Close
Set Conn=Nothing
%>
</body>
</html>