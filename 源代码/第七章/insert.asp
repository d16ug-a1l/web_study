<!--#include file="adovbs.inc"-->
<%
'创建ADODB.Connection对象
Set Conn=Server.Createobject("Adodb.Connection") 

'依据连接的数据库设置连接字符串
Conn.ConnectionString="Provider=Microsoft.Jet.OLEDB.4.0;"&_
			 			"Data Source="&Server.MapPath("oa.mdb")
Conn.Open  '打开与数据库的连接
Set rs=Server.CreateObject("ADODB.Recordset")
rs.Open "Group_Info",Conn,adOpenKeyset,adLockOptimistic,adCmdTable
Name=Request.Form("MingCheng")
Info=Request.Form("XinXi")
Response.write Name&Info
rs.Addnew Array("Name","Info"),Array(Name,Info)
rs.Update
rs.close
Set rs=Nothing
Conn.Close
Set Conn=Nothing
%>
