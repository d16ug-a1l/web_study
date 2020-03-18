<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>新建网页 1</title>
</head>

<body>
<%
Set Conn=Server.Createobject("Adodb.Connection") 
Conn.ConnectionString="Provider=Microsoft.Jet.OLEDB.4.0;"&_
			 			"Data Source="&Server.MapPath("/oa.mdb")
Conn.Open  '打开与数据库的连接
  Set rs = conn.OpenSchema(adSchemaTables,TABLE_NAME)
   Do While Not rs.EOF 
      Response.Write "表的名称："&rs("TABLE_NAME")&";表的类型为："&rs("TABLE_TYPE")&"<br />"
      rs.MoveNext 
   Loop 
   rs.Close
%>
</body>

</html>
