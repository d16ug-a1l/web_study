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
      Response.Write "<option value='"&rs("TABLE_NAME")&";"&rs("TABLE_TYPE")&"' >"&rs("TABLE_NAME")&"��������ͣ�"&rs("TABLE_TYPE")&"</option><br />"
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
	Response.write "�������Ϊ��<font color='#FF0000'>"&TableName&"</FONT>��  �������Ϊ��<font color='#FF0000'>"&TableType&"</FONT><BR>"
	If TableName="" or TableType="" Then
		Response.write "<font color='#FF0000'>�������߱�����Ϊ�գ�</FONT>"
		Response.end
	End If
	If Instr(TableType,"ACCESS")>0 or Instr(TableType,"SYSTEM")>0 Then
		Response.write "<font color='#FF0000'>����ϵͳ�ļ���û�ж�ȡȨ��</FONT>"
		Response.end
	End If
	sql="select * from "&TableName
 
	'ʹ��state�����жϵ�ǰ���ӵ�״̬
	Set rs=Conn.execute(sql)
	For i=0 to rs.fields.Count-1
		Response.write rs.Fields(i).Name&":"
		Response.write GetInfo(i)&"<BR>"
	Next
	If Conn.State<>1 Then
	  Response.write "<font color='#FF0000'>�����ݿ⽨�������Ӳ��ɹ���"
	  Response.end
	End If
	
	Conn.Close
	Set Conn=Nothing
	Function GetInfo(i)
	   Select Case rs.Fields(i).Type
	      Case 202
	        str="�ַ����͡���ֵΪ��<font color='#FF0000'>"&rs.Fields(i).Value&"</FONT><BR>"
	      Case 203
	        str="��ע�͡���ֵΪ��<font color='#FF0000'>"&rs.Fields(i).Value&"</FONT><BR>"
	      Case 3
	         str="�Զ�����͡���ֵΪ��<font color='#FF0000'>"&rs.Fields(i).Value&"</FONT>����ֵ��Χ�ǣ�<font color='#FF0000'>"&rs.Fields(i).NumericScale&"</FONT><BR>"
	      Case 2
	        str="���͡���ֵΪ��<font color='#FF0000'>"&rs.Fields(i).Value&"����ֵ��Χ�ǣ�<font color='#FF0000'>"&rs.Fields(i).NumericScale&"</FONT><BR>"
	      Case 7
	        str="����ʱ���͡���ֵΪ��<font color='#FF0000'>"&rs.Fields(i).Value&"����ֵ��Χ�ǣ�<font color='#FF0000'>"&rs.Fields(i).NumericScale&"</FONT><BR>"
	   End Select
		GetInfo=str
	End Function
End If
%>
</body>
</html>