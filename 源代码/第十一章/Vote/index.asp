<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>����ͶƱϵͳ</title>
</head>
<body>
<div align="center">
<font face="�����п�" size="6" color="#0000FF">����ͶƱϵͳ</font> 
 <form method="post" action="vote.asp">
<table border="1" width="80%">
<tr>
 <td bgcolor="#C0C0C0" align="center" height="21" >ͶƱ��Ŀ</td>
</tr>
<%
 Set Conn=Server.Createobject("Adodb.Connection") 
 Conn.ConnectionString="Provider=Microsoft.Jet.OLEDB.4.0;"&_
  			"Data Source="&Server.MapPath("vote.mdb")
 Conn.Open
 Set rs=Server.Createobject("Adodb.Recordset") 
 Sql="Select * from VoteItem where IsDisp='1'" 
 
 rs.Open Sql,Conn,1,1
do while rs.EOF=False
 
 Response.write("<tr><td><input Type='CHECKBOX' NAME='checkbox' VALUE='"&rs("Item")&"'>"&rs("Item")&"</td></tr>")
 
 rs.movenext 
loop 
%>
</table> 
<p bgcolor="#C0C0C0" align=center>
  <input type="submit" name="T1" value='ͶƱ' size="28"  > 
  <input type="BUTTON" name="T2" value='�鿴ͶƱ���' size="28"
      OnClick="vbscript:Window.open 'result.asp','_self'" >
</p> 
</form>
</div>
</body>

</html>