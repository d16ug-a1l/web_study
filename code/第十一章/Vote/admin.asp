<html>

<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 6.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>����ͶƱϵͳ</title>
</head>

<body>
<p align="center">
<font face="�����п�" size="6" color="#0000FF">����ͶƱ�������</font></p>
<div align="center">
<table border="1" width="60%" align=center>
  <tr>
    <td width="30%" align="center">��Ŀ</td>
    <td width="25%" align="center">�޸�</td>
    <td width="25%" align="center">ɾ��</td>
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
str="<tr><td width='30%'>"&rs("Item")&"</td>"&_
    "<td width='25%' align='center'>"&"<a href='modify.asp?ID="&rs("ID")&"&action=modify'>�޸�</a></td>"&_
	"<td width='25%' align='center'><a href='modify.asp?ID="&rs("ID")&"&action=del'>ɾ��</a></td></tr>" 
	Response.write(str)
	rs.movenext 
loop 
%>
</table>
<form method=post action="modify_1.asp?action=add">
<table border=0  width="60%" >
  <tr>
    <td width="49%" valign="middle" align="left" bordercolor="#000000">ͶƱ��Ŀ��</td>
    <td width="69%" valign="middle" align="left" bordercolor="#000000">
    	<input type="text" name="Content"   size="22">
    </td>
    <td> <input type="submit" name="T2" value='������Ŀ' size="28" ></td>
  </tr>
  </table>
 </form>
</div>
</body>

</html>