<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>UPDATE����������</title>
</head>
<body>
<%
   dim conn
   dim sql
   Set conn = Server.CreateObject("ADODB.Connection")
   conn.Open "driver={Microsoft Access Driver (*.mdb)};dbq="&Server.MapPath("db_student.mdb")
   'UPDATE����������
   sql="UPDATE T_STUDENT SET T_S_NAME='�ž�'"
   sql=sql&"WHERE T_S_ID =2002080531"
   conn.execute sql
   conn.close 
   set conn=nothing 
 %>
 <script language="vbscript">
   MsgBox "����T_STUDENT����һ������", , "�ɹ���ʾ"
 </script> 
</body>
</html>
