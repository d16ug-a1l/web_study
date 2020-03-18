<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 6.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>New Page 1</title>
</head>

<body>
<%
Dim strContent 
Dim Action,strID
Action=Request.QueryString("action")
strID=Request.QueryString("ID")
'获取节点各个字段值
strContent=Request.Form("Content")
Set Conn=Server.Createobject("Adodb.Connection") 
Conn.ConnectionString="Provider=Microsoft.Jet.OLEDB.4.0;"&_
  			"Data Source="&Server.MapPath("vote.mdb")
 Conn.Open
If Action="add" Then
   Set rs=Server.Createobject("Adodb.Recordset") 
   Sql="insert into VoteItem(Item) values('"&strContent&"')" 
   Conn.Execute(Sql)
   Conn.Close
   Response.write(strContent&"项目添加成功")
ElseIF Action="modify" Then
   Sql="update VoteItem set [Item]='"&strContent&"' where ID="&strID 
   Conn.Execute(Sql)
   Response.write(strContent&"项目修改成功")
End If

%>
</body>

</html>

