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
 			"Data Source="&Server.MapPath("vote.mdb")
Conn.Open
Set rs=Server.Createobject("Adodb.Recordset") 
Sql="Select * from VoteItem where ID="&ActionID 
If Action="add" Then
   strAct="����"
ElseIf Action="modify" Then
  strAct="�޸�"
  rs.Open Sql,Conn,1,1
  strContent=rs("Item")
ElseIf Action="del" Then
 Sql="delete from VoteItem where ID="&ActionID 
 Conn.Execute(Sql)
 Response.write("��Ŀ�ɹ�ɾ��")
End If

if Action="modify" or Action="add" Then
%>
<form method="POST" action=<% =response.Write("Modify_1.asp?action="&Action&"&ID="&ActionID) %> >
<p align="center"><% =strAct %>�ּ�Ŀ¼��Ŀ</p>
<table border="0" width="39%" align="center">
  <tr>
    <td width="49%" valign="middle" align="left" bordercolor="#000000">ͶƱ��Ŀ��</td>
    <td width="69%" valign="middle" align="left" bordercolor="#000000">
    	<input type="text" name="Content" value=<% =response.write("'"&strContent&"'") %> size="22">
    </td>
  </tr>
</table>
  <p align="center"><input type="submit" value='<% =strAct %>' name="B1"><input type="reset" value="ȫ����д" name="B2"></p>
</form>
<%
end if
%>
</body>

</html>