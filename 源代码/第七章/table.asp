<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�½���ҳ 1</title>
</head>

<body>
<%
Set Conn=Server.Createobject("Adodb.Connection") 
Conn.ConnectionString="Provider=Microsoft.Jet.OLEDB.4.0;"&_
			 			"Data Source="&Server.MapPath("/oa.mdb")
Conn.Open  '�������ݿ������
  Set rs = conn.OpenSchema(adSchemaTables,TABLE_NAME)
   Do While Not rs.EOF 
      Response.Write "������ƣ�"&rs("TABLE_NAME")&";�������Ϊ��"&rs("TABLE_TYPE")&"<br />"
      rs.MoveNext 
   Loop 
   rs.Close
%>
</body>

</html>
