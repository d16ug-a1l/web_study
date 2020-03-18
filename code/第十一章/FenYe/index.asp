
<%
Set Conn=Server.Createobject("Adodb.Connection") 
Conn.ConnectionString="Provider=Microsoft.Jet.OLEDB.4.0;"&_
 			"Data Source="&Server.MapPath("user.mdb")
Conn.Open
Set rs=Server.Createobject("Adodb.Recordset") 
Sql="Select * from LayerST "
rs.Open Sql,Conn,1,1

Dim nPageSize,nPSize,nPageNo,nPageCount
nPageSize=2

If rs.RecordCount=0 Then
	Response.Write("没有指定的记录！")
Else
	rs.PageSize = Cint(nPageSize) 
	nPageNo=CInt(Request.QueryString("page"))
	nPageCount=rs.PageCount
	If nPageNo<1 Then
		nPageNo=1
	ElseIf nPageNo>rs.PageCount Then
		nPageNo=rs.PageCount
	End If
	rs.AbsolutePage=nPageNo 
	%>
<p align="center"><font face="华文行楷" size="6" color="#0000FF">分页显示</font></p>
<div align="center">
<table width="200" border="1" >
  <tr>
    <td align="center">序号</td>
    <td align="center">内容</td>
    <td align="center">代码</td>
</tr>
  <%
	For i=1 to rs.pagesize
		Response.Write("<TR><td>"&rs("ID")&"</td><td>"&_
		rs("Title")&"</td><td>"&rs("Layer")&"</td></tr>")
		rs.movenext
		If rs.EOF Then Exit For
	Next
	Response.Write("</table>")
	rs.Close
	Conn.Close
End If
If nPageNo=1 Then
	Response.Write("首页")
ElseIf nPageNo>1 Then
	Response.Write("<a href=index.asp?page="&nPageNo-1&">前一页</a>")
End If
If nPageNo=nPageCount Then
	Response.Write("尾页")
ElseIf nPageNo<nPageCount Then
	Response.Write("<a href=index.asp?page="&nPageNo+1&">下一页</a>")
End If
Response.write("  </div>")
%>