<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 6.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>New Page 1</title>
</head>

<body>
<%
Dim Action
Dim ActionID
Dim strAct
Dim strTitle,strContent,strType,strLayer,strIsDel
Action=Request.QueryString("action")
ActionID=Request.QueryString("ID")
Set Conn=Server.Createobject("Adodb.Connection") 
Conn.ConnectionString="Provider=Microsoft.Jet.OLEDB.4.0;"&_
 			"Data Source="&Server.MapPath("user.mdb")
Conn.Open
Set rs=Server.Createobject("Adodb.Recordset") 
Sql="Select * from LayerST where ID="&ActionID 
rs.Open Sql,Conn,1,1
If Action="add" Then
   strAct="增加"
   strLayer=rs("Layer")
ElseIf Action="modify" Then
  strAct="修改"
  rs.Open Sql,Conn,1,1
  strTitle=rs("Title")
  strContent=rs("Content")
  strType=rs("Type")
  strIsDel=rs("IsDisp")	
ElseIf Action="del" Then
 Sql="delete from LayerST where Layer Like '"&rs("Layer")&"%'"
 Conn.Execute(Sql)
 Response.write("项目成功删除")
End If

if Action="modify" or Action="add" Then
%>
<form method="POST" action=<% =response.Write("Modify_1.asp?action="&Action&"&ID="&ActionID) %> >
 
<p align="center"><% =strAct %>分级目录项目</p>
<table border="0" width="39%" align="center">
  <tr>
    <td width="49%" valign="middle" align="left" bordercolor="#000000">标题：</td>
    <td width="69%" valign="middle" align="left" bordercolor="#000000">
    	<input type="text" name="Title" value=<% =response.write("'"&strTitle&"'") %> size="22">
    </td>
  </tr>
  <tr>
    <td width="49%" valign="middle" align="left" bordercolor="#000000">内容：</td>
    <td width="69%" valign="middle" align="left" bordercolor="#000000">
       <input type="text" name="Title" value=<% =response.write("'"&strContent&"'") %> size="22">
    </td>
  </tr>
  <%
  If Action="add" Then
  %>
  <tr>
    <td width="49%" valign="middle" align="left" bordercolor="#000000">层：</td>
    <td width="69%" valign="middle" align="left" bordercolor="#000000">
    	<input type="text" name="Layer" value=<% =response.write("'"&strLayer&"'") %> size="22"> 
    </td>
  </tr>
  <%
  End if
  %>
  <tr>
    <td width="49%" valign="middle" align="left" bordercolor="#000000">类型：</td>
    <td width="69%" valign="middle" align="left" bordercolor="#000000">
    	<input type="text" name="Layer" value=<% =response.write("'"&strType&"'") %> size="22">
    </td>
  </tr>
  <tr>
    <td width="49%" valign="middle" align="left" bordercolor="#000000">是否显示：</td>
    <td width="69%" valign="middle" align="left" bordercolor="#000000">
        <input type="text" name="Layer" value=<% =response.write("'"&strIsDel&"'") %> size="22">
    </td>
  </tr>
</table>
  <p align="center"><input type="submit" value='<% =strAct %>' name="B1"><input type="reset" value="全部重写" name="B2"></p>
</form>
<%
end if
%>
</body>

</html>
