<!--#include file="adovbs.inc"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title></title>
</head>


<body>
<script language="vbscript">
Sub ReDirectTable()
	location.href="13.5.2.asp?TableName="&window.Userform.TableName.value
End Sub
Sub ReDirectColumn()
	location.href="13.5.2.asp?TableName="&window.Userform.TableName.value&"&ColumnName="&window.Userform.ColumnName.value
End Sub
</script>
<%
'����ADO DB.Connection����
Set Conn=Server.Createobject("Adodb.Connection") 

'�������ӵ����ݿ����������ַ���
Conn.ConnectionString="Provider=Microsoft.Jet.OLEDB.4.0;"&_
			 			"Data Source="&Server.MapPath("oa.mdb")
Conn.Open  '�������ݿ������
Name=Request.QueryString("TableName")
ColumnName=Trim(Request.QueryString("ColumnName"))
%>
<p style="margin-top: 0; margin-bottom: 0" align="center">
<font size="6" color="#0000FF" face="�����п�">��ȡ�ֶε�ֵ</font></p>
<p style="margin-top: 0; margin-bottom: 0">
<form name="Userform">
<table align=center>
<tr><td>
������ƣ�<select name="TableName" onchange="ReDirectTable()">
<%
   Set rs = conn.OpenSchema(adSchemaTables,TABLE_NAME)
   Do While Not rs.EOF 
     If not(Instr(rs("TABLE_TYPE"),"ACCESS")<>0 or  Instr(rs("TABLE_TYPE"),"SYSTEM")<>0) Then
     	If rs("TABLE_NAME")=Name  Then
	     	Response.Write "<option value='"&rs("TABLE_NAME")&"' selected>"&rs("TABLE_NAME")&"</option><br/>"
	     Else
      		Response.Write "<option value='"&rs("TABLE_NAME")&"' >"&rs("TABLE_NAME")&"</option><br/>"
      	End If
     End If
      rs.MoveNext 
   Loop 
   rs.Close
%>
</select>
</td></tr>
<tr><td>
�ֶε����ƣ�<select name="ColumnName" onchange="ReDirectColumn()">
<%

If Trim(Name)<>"" Then
	Set rs = conn.OpenSchema(adSchemaColumns,COLUMN_NAME)
    Do While Not rs.EOF 
    	'Response.write rs("TABLE_NAME")
    	If Trim(rs("TABLE_NAME"))=Name Then
        	Response.Write "<option value='"&rs("COLUMN_NAME")&"' >"&rs("COLUMN_NAME")&"</option><br/>"
        End If
      rs.MoveNext 
   Loop 
   rs.Close
End If
%>
</select>
<%
REsponse.write "<BR>ѡ�еı�Ϊ��"&Name
	If ColumnName<>"" and Name<>"" Then
		sql="select "&ColumnName&" from "&Name
		Set rs=Conn.execute(sql)
		Dim i
		i=1
		REsponse.write "<BR>��ȡ���ֶ��ǣ�"&ColumnName&"��ֵΪ��<BR>"
		Response.write "<table border=1 align=center>"
		Do while not rs.Eof
			If i mod 3 =1 Then Response.write "<tr>"
			Response.write "<td>"&rs(ColumnName).Value&"</td>"
			If i mod 3=0 Then Response.write "</tr>"
			i=i+1
			rs.movenext
		Loop
	End If
	Conn.Close
	Set Conn=Nothing
%>
</td></tr>
</p>
</table>
</form>
</body>

</html>
