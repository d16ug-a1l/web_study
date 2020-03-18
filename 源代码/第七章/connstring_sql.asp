<!--#include file="adovbs.inc"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title></title>
</head>


<body>
<script language="vbscript">
Sub ReDirectLocal()
	location.href="connstring_sql.asp?TableName="&window.Userform.TableName.value
End Sub
</script>
<%
Set Conn= Server.CreateObject("ADODB.Connection")
Conn.Open "driver={SQL Server};server=(Local);uid=sa;pwd=;database=ASP"
 
%>
<form name="Userform">
<select name="TableName" onchange="ReDirectLocal()">
<%
   Set rs = Conn.OpenSchema(adSchemaTables,TABLE_NAME)
   Do While Not rs.EOF 
      Response.Write "<option value='"&rs("TABLE_NAME")&";"&rs("TABLE_TYPE")&"' >"&rs("TABLE_NAME")&"。表的类型："&rs("TABLE_TYPE")&"</option><br />"
      rs.MoveNext 
   Loop 
   rs.Close
%>
</select>
</form>
<%
Name=Request.QueryString("TableName")

If Trim(Name)<>"" Then
	Table=Split(Name,";")
	TableName=Trim(Table(0))
	TableType=Trim(Table(1))
	Response.write "表的名称为：<font color='#FF0000'>"&TableName&"</FONT>。  表的类型为：<font color='#FF0000'>"&TableType&"</FONT><BR>"
	If TableName="" or TableType="" Then
		Response.write "<font color='#FF0000'>表名或者表类型为空！</FONT>"
		Response.end
	End If
	If Instr(TableType,"ACCESS")>0 or Instr(TableType,"SYSTEM")>0 Then
		Response.write "<font color='#FF0000'>这是系统文件，没有读取权限</FONT>"
		Response.end
	End If
	sql="select * from "&TableName
 
	'使用state属性判断当前连接的状态
	Set rs=Conn.execute(sql)
	For i=0 to rs.fields.Count-1
		Response.write rs.Fields(i).Name&":"
		Response.write GetInfo(i)&"<BR>"
	Next
	If Conn.State<>1 Then
	  Response.write "<font color='#FF0000'>与数据库建立的连接不成功！"
	  Response.end
	End If
	
	Conn.Close
	Set Conn=Nothing
	Function GetInfo(i)
	   Select Case rs.Fields(i).Type
	      Case 202
	        str="字符串型。其值为：<font color='#FF0000'>"&rs.Fields(i).Value&"</FONT><BR>"
	      Case 203
	        str="备注型。其值为：<font color='#FF0000'>"&rs.Fields(i).Value&"</FONT><BR>"
	      Case 3
	         str="自动编号型。其值为：<font color='#FF0000'>"&rs.Fields(i).Value&"</FONT>。数值范围是：<font color='#FF0000'>"&rs.Fields(i).NumericScale&"</FONT><BR>"
	      Case 2
	        str="整型。其值为：<font color='#FF0000'>"&rs.Fields(i).Value&"。数值范围是：<font color='#FF0000'>"&rs.Fields(i).NumericScale&"</FONT><BR>"
	      Case 7
	        str="日期时间型。其值为：<font color='#FF0000'>"&rs.Fields(i).Value&"。数值范围是：<font color='#FF0000'>"&rs.Fields(i).NumericScale&"</FONT><BR>"
	   End Select
		GetInfo=str
	End Function
End If
%>
</body>
</html>