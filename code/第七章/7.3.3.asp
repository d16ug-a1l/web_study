<!--#include file="adovbs.inc"-->
<html>

<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�½���ҳ 1</title>
</head>

<body>
<table ><tr><td>
<form method="POST" action="insert.asp">
	<p>ְλ���ƣ�<input type="text" name="MingCheng" size="20"></p>
	<p>ְλ������<input type="text" name="XinXi" size="20"></p>
	<p><input type="submit" value="�ύ" name="B1"><input type="reset" value="����" name="B2"></p>
</form>
<%
'����ADODB.Connection����
Set Conn=Server.Createobject("Adodb.Connection") 

'�������ӵ����ݿ����������ַ���
Conn.ConnectionString="Provider=Microsoft.Jet.OLEDB.4.0;"&_
			 			"Data Source="&Server.MapPath("oa.mdb")
Conn.Open  '�������ݿ������
Set rs=Server.CreateObject("ADODB.Recordset")
rs.Open "Group_Info",Conn,adOpenKeyset,adLockOptimistic,adCmdTable
Do While not rs.EOF 
	Response.write rs("Name")&"  <a href='delete.asp?id="&rs("ID")&"'>ɾ��</a><BR>"
	rs.movenext
Loop
rs.close
Set rs=Nothing
Conn.Close
Set Conn=Nothing
%>
</td></tr></table>
</body>

</html>