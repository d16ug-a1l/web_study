<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>在线投票系统</title>
</head>

<body>
<div align="center">
<font face="华文行楷" size="6" color="#0000FF">投票结果</font> 
 <form method="post" action="vote.asp">
<table border="1" width="60%">
<tr>
 <td bgcolor="#C0C0C0" align="center" height="21"  >投票项目</td>
 <td bgcolor="#C0C0C0" align="center" height="21"  >得票数</td>
 <td bgcolor="#C0C0C0" align="center" height="21"  >得票百分比</td>
</tr>

<%
 Set Conn=Server.Createobject("Adodb.Connection") 
 Conn.ConnectionString="Provider=Microsoft.Jet.OLEDB.4.0;"&_
  			"Data Source="&Server.MapPath("vote.mdb")
 Conn.Open
 Set rs=Server.Createobject("Adodb.Recordset") 
 Sql="Select Sum(Count) As Total from VoteItem where IsDisp='1'" 
 
 Set rs=Conn.Execute(Sql)
 If rs.EOF=False Then nTotal=rs("Total")
  Sql="Select * from VoteItem where IsDisp='1'" 
 Dim percent
 Set rs=Conn.Execute(Sql)
do while rs.EOF=False
 If nTotal=0 Then
 	percent=0
 Else
 	percent=(rs("Count")/nTotal)*100
 	percent=formatNumber(percent,1)
 End If
 Response.write("<tr>") 
 Response.write("<td>"&rs("Item")&"</td> ")
 Response.write("<td>"&rs("Count")&"</td> ")
 Response.write("<td><img src='img/bar.gif' width='"&percent*2&"' height=20>" &percent&"%</td></tr> ")
 rs.movenext 
loop 
%>
<tr>
	<td align="left">
<%
	Response.write("总票数："&nTotal)
%>
	</td>
</tr>
</table>
</form>
</div>
</body>

</html>