<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 6.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>New Page 1</title>
</head>

<body>
<%
Dim strTitle,strContent,strType,strLayer,strIsDel
Dim Action,strID
Action=Request.QueryString("action")
strID=Request.QueryString("ID")
'获取节点各个字段值
strTitle=Request.Form("Title")
strContent=Request.Form("Content")
strType=Request.Form("Type")
strIsDel=Request.Form("IsDisp")
 Set Conn=Server.Createobject("Adodb.Connection") 
 Conn.ConnectionString="Provider=Microsoft.Jet.OLEDB.4.0;"&_
  			"Data Source="&Server.MapPath("user.mdb")
 Conn.Open
If Action="add" Then
   strLayer=Request.Form("Layer")
   strLayer=Mid(strLayer,1,2)
   Set rs=Server.Createobject("Adodb.Recordset") 
   Sql="Select * from LayerST where ID="&strID 
   ' Response.write(Sql)
   rs.Open Sql,Conn,1,1
   strLayer=rs("Layer")&strLayer
   rs.close
   Sql="insert into LayerST(TiTle,Content,Layer,Type,IsDisp) values('"&_
   strTitle&"','"&strContent&"','"&strLayer&"','"&strType&"','"&strIsDel&"')" 
   Conn.Execute(Sql)
   Conn.Close
   Response.write("项目添加成功")
ElseIF Action="modify" Then
   Sql="update [LayerST] set [Title]='"&strTitle&_
   "', [Content]='"&strContent&"', [Type]='"&strType&_
   "', [IsDisp]='"&strIsDel&"' where ID="&strID 
 
   Conn.Execute(Sql)
   Response.write("项目修改成功")
End If

%>
</body>

</html>
