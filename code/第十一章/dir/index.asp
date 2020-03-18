<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>分级目录模块</title>
</head>
 <script>
function showObj(str,imgid) {
  divObj=eval(str);
  imgObj=eval(imgid);
  if (divObj.style.display=="none") {
    imgObj.src="img/open.gif";
    divObj.style.display="inline";
  }
  else {
    imgObj.src="img/plus.gif";
    divObj.style.display="none";
  }

}
 </script>
<body>
<%
 call init()

Sub init()
	call Generate("__")
End Sub 
 '生成文档节点
Sub WriteDoc(layer,title,Content)
For i=1 To len(layer) 
	Response.Write("&nbsp;")
Next 
Response.write("<a href="&Content&" >"&title&"</a><br>")
End Sub 

 '生成文档节点
Sub WriteNode (layer,title)
For i=1 To len(layer) 
	Response.Write("&nbsp;")
Next 
Response.write("<img id=img"&layer&" src='img/plus.gif' border=0>"&_
			   "<a href='#' onclick=showObj('id"&layer&"','img"&layer&_
			   "')>"&title&" </a><br>")
Response.write("<div id=id"&layer&" style='display:none'>")

 
End Sub 

Sub Generate (layer)
  Dim parent
  If layer="__" Then
    parent="__"
  Else
    parent=layer
  End If
 Set Conn=Server.Createobject("Adodb.Connection") 
 Conn.ConnectionString="Provider=Microsoft.Jet.OLEDB.4.0;"&_
  			"Data Source="&Server.MapPath("user.mdb")
 Conn.Open
 Set rs=Server.Createobject("Adodb.Recordset") 
 Sql="Select * from LayerST where IsDisp='1' and Layer Like '"&layer&"'" 
 
 rs.Open Sql,Conn,1,1
 do while rs.EOF=False
   If rs("Type")="1" Then 
     call WriteNode(rs("Layer"),rs("Title"))
     call Generate (rs("Layer")&"__")
	 Response.write("</div>")
   Else 
     call WriteDoc(rs("Layer"),rs("Title"),rs("Content") )	
   End If
  rs.movenext 
loop 
 
rs.close
Conn.close

End Sub 
%>
</body>

</html>
