<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>UPDATE语句更新数据</title>
</head>
<body>
<%
   dim conn
   dim sql
   Set conn = Server.CreateObject("ADODB.Connection")
   conn.Open "driver={Microsoft Access Driver (*.mdb)};dbq="&Server.MapPath("db_student.mdb")
   'UPDATE语句更新数据
   sql="UPDATE T_STUDENT SET T_S_NAME='张竟'"
   sql=sql&"WHERE T_S_ID =2002080531"
   conn.execute sql
   conn.close 
   set conn=nothing 
 %>
 <script language="vbscript">
   MsgBox "更新T_STUDENT表中一条数据", , "成功提示"
 </script> 
</body>
</html>
