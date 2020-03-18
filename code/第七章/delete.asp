<!--#include file="adovbs.inc"-->
<%
'创建ADODB.Connection对象
Set Conn=Server.Createobject("Adodb.Connection") 

'依据连接的数据库设置连接字符串
Conn.ConnectionString="Provider=Microsoft.Jet.OLEDB.4.0;"&_
			 			"Data Source="&Server.MapPath("oa.mdb")
Conn.Open  '打开与数据库的连接
Set rs=Server.CreateObject("ADODB.Recordset")
ID1=Request.QueryString("ID")
Dim sql
sql="SELECT * FROM [Group_Info] WHERE  [ID]="&ID1
Response.write sql
rs.Open sql,Conn,adOpenKeyset,,adCmdTable
rs.Delete
rs.Update
rs.close
Set rs=Nothing
Conn.Close
Set Conn=Nothing
%>
