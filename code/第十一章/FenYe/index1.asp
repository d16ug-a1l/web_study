<%
'连接数据库
Set Conn=Server.Createobject("Adodb.Connection") 
Conn.ConnectionString="Provider=Microsoft.Jet.OLEDB.4.0;"&_
 			"Data Source="&Server.MapPath("user.mdb")
Conn.Open
Sql="Select count(*) As RecordCount from LayerST"
Set rs=Conn.Execute(Sql)
Dim nPageSize,nPageCount,nCursePos,nCount
nCount=rs("RecordCount")
nPageSize=2
nPageCount=Int((nCount/nPageSize)*(-1))*(-1)
nPageNo=CInt(Request.QueryString("page"))
If nPageNo<1 Then
	nPageNo=1
ElseIf nPageNo>nPageCount Then
	nPageNo=nPageCount
End If
If Request.QueryString("CurseID")="" Then 
	nCursePos=0 
Else 
	nCursePos=Clng(Request.QueryString("CurseID"))
End If
If Request.QueryString("Type")="" Then 
	strType="next" 
Else 
	strType=Request.QueryString("Type")
End If
If strType="next" Then
	Sql="Select Top "&nPageSize&" ID,Content,Title,Layer From LayerST where ID>"&nCursePos
Else
	Sql="Select Top "&nPageSize&" ID,Content,Title,Layer From LayerST where ID IN"&_
		" (Select Top "&nPageSize&" ID From LayerST where ID<"&nCursePos&_
		" order by ID DESC) order by ID "
End IF
%>
<p align="center"><font face="华文行楷" size="6" color="#0000FF">分页显示</font></p>
<div align="center">
<table width="200" border="1" >
  <tr>
    <td align="center">标题</td>
    <td align="center">内容</td>
    <td align="center">代码</td>
  </tr>


<%
Dim nCurseStart,nCurseEnd

Set rs= Conn.Execute(Sql)
For i=1 to nPageSize
	If rs.EOF Then Exit For
	If i=1 Then nCurseStart=rs.Fields("ID")
	Response.Write("<TR><td>"&rs.Fields("Title")&"</td><td>"&rs.Fields("ID")&"</td><td>"&rs.Fields("Layer")&"</td></tr>")
	nCurseEnd=rs.Fields("ID")
	rs.movenext
Next

Response.Write("</table> ")
Conn.Close
If nPageNo=1 Then
	Response.Write("首页")
ElseIf nPageNo>1 Then
	Response.Write("<a href=index1.asp?type=before&page="&nPageNo-1&"&CurseID="&nCurseStart&">前一页</a>")
End If
If nPageNo=nPageCount Then
	Response.Write("尾页")
ElseIf nPageNo<nPageCount Then
	Response.Write("<a href=index1.asp?type=next&page="&nPageNo+1&"&CurseID="&nCurseEnd&">下一页</a>")
End If
Response.Write(" </div>")
%>