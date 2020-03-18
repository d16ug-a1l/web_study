<html>

<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>New Page 1</title>
</head>

<body>
<table border="1" width="100%">
  <tr>
    <td width="25%" align="center">项目(层)</td>
    <td width="25%" align="center">修改</td>
    <td width="25%" align="center">增加</td>
    <td width="25%" align="center">删除</td>
  </tr>
<%
 Set Conn=Server.Createobject("Adodb.Connection") 
 Conn.ConnectionString="Provider=Microsoft.Jet.OLEDB.4.0;"&_
  			"Data Source="&Server.MapPath("user.mdb")
 Conn.Open
 Set rs=Server.Createobject("Adodb.Recordset") 
Sql="Select * from LayerST order by Layer " 
rs.Open Sql,Conn,1,1
 
do while  rs.EOF=False 
str="<tr><td width='25%'>"&rs("Title")&"("&rs("Layer")&") </td>"&_
    "<td width='25%' align='center'>"&"<a href='modify.asp?ID="&rs("ID")&"&action=modify'>修改</a>"&_
	"</td><td width='25%' align='center'><a href='modify.asp?ID="&rs("ID")&"&action=add'>增加</a>"&_
	"</td><td width='25%' align='center'><a href='modify.asp?ID="&rs("ID")&"&action=del'>删除</a></td></tr>"
Response.write(str)
rs.movenext
Loop
rs.close
Conn.close
%>

</table>
</body>

</html>

