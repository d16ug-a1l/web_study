<!--#include file="adovbs.inc"-->
<html>

<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>新建网页 1</title>
</head>

<body>
<table ><tr><td>
<form method="POST" action="insert.asp">
	<p>职位名称：<input type="text" name="MingCheng" size="20"></p>
	<p>职位描述：<input type="text" name="XinXi" size="20"></p>
	<p><input type="submit" value="提交" name="B1"><input type="reset" value="重置" name="B2"></p>
</form>
<%
'创建ADODB.Connection对象
Set Conn=Server.Createobject("Adodb.Connection") 

'依据连接的数据库设置连接字符串
Conn.ConnectionString="Provider=Microsoft.Jet.OLEDB.4.0;"&_
			 			"Data Source="&Server.MapPath("oa.mdb")
Conn.Open  '打开与数据库的连接
Set rs=Server.CreateObject("ADODB.Recordset")
rs.Open "Group_Info",Conn,adOpenKeyset,adLockOptimistic,adCmdTable
Do While not rs.EOF 
	Response.write rs("Name")&"  <a href='delete.asp?id="&rs("ID")&"'>删除</a><BR>"
	rs.movenext
Loop
rs.close
Set rs=Nothing
Conn.Close
Set Conn=Nothing
%>
</td></tr></table>
</body>

</html>