<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>����ͶƱϵͳ</title>
</head>

<body>
<div align="center">
<font face="�����п�" size="6" color="#0000FF">ͶƱ���</font> 
 <form method="post" action="vote.asp">
<table border="1" width="60%">
<tr>
 <td bgcolor="#C0C0C0" align="center" height="21"  >ͶƱ��Ŀ</td>
 <td bgcolor="#C0C0C0" align="center" height="21"  >��Ʊ��</td>
 <td bgcolor="#C0C0C0" align="center" height="21"  >��Ʊ�ٷֱ�</td>
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
	Response.write("��Ʊ����"&nTotal)
%>
	</td>
</tr>
</table>
</form>
</div>
</body>

</html>