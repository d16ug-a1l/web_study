<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>在线投票系统</title>
</head>
<body>
<div align="center">
<font face="华文行楷" size="6" color="#0000FF">在线投票系统</font> 
 <form method="post" action="vote.asp">
<table border="1" width="80%">
<tr>
 <td bgcolor="#C0C0C0" align="center" height="21" >投票项目</td>
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
  <input type="submit" name="T1" value='投票' size="28"  > 
  <input type="BUTTON" name="T2" value='查看投票结果' size="28"
      OnClick="vbscript:Window.open 'result.asp','_self'" >
</p> 
</form>
</div>
</body>

</html>