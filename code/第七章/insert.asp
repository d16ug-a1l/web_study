<!--#include file="adovbs.inc"-->
<%
'����ADODB.Connection����
Set Conn=Server.Createobject("Adodb.Connection") 

'�������ӵ����ݿ����������ַ���
Conn.ConnectionString="Provider=Microsoft.Jet.OLEDB.4.0;"&_
			 			"Data Source="&Server.MapPath("oa.mdb")
Conn.Open  '�������ݿ������
Set rs=Server.CreateObject("ADODB.Recordset")
rs.Open "Group_Info",Conn,adOpenKeyset,adLockOptimistic,adCmdTable
Name=Request.Form("MingCheng")
Info=Request.Form("XinXi")
Response.write Name&Info
rs.Addnew Array("Name","Info"),Array(Name,Info)
rs.Update
rs.close
Set rs=Nothing
Conn.Close
Set Conn=Nothing
%>
