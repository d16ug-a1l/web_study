<!--#include file="adovbs.inc"-->
<%
'����ADODB.Connection����
Set Conn=Server.Createobject("Adodb.Connection") 

'�������ӵ����ݿ����������ַ���
Conn.ConnectionString="Provider=Microsoft.Jet.OLEDB.4.0;"&_
			 			"Data Source="&Server.MapPath("oa.mdb")
Conn.Open  '�������ݿ������
Set rs=Server.CreateObject("ADODB.Recordset")
ID1=Request.QueryString("ID")
Dim sql
sql="SELECT * FROM [Group_Info] WHERE  [ID]="&ID1
Response.write sql
rs.Open sql,Conn,adOpenKeyset,,adCmdTable
rs.Delete
rs.Update
rs.close
Set rs=Nothing
Conn.Close
Set Conn=Nothing
%>
