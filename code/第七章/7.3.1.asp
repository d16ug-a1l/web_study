<!--#include file="adovbs.inc"-->
<html>

<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�½���ҳ 1</title>
</head>

<body>
<%
'����ADO DB.Connection����
Set Conn=Server.Createobject("Adodb.Connection") 

'�������ӵ����ݿ����������ַ���
Conn.ConnectionString="Provider=Microsoft.Jet.OLEDB.4.0;"&_
			 			"Data Source="&Server.MapPath("oa.mdb")
Conn.Open  '�������ݿ������
Set rs=Server.CreateObject("ADODB.Recordset")
rs.Open "Group_Info",Conn,adOpenKeyset,adLockOptimistic,adCmdTable
If not rs.EOF then
	Do while not rs.Eof
		Response.write "<font color='#FF0000'>ְλ��</font>"&rs("Name")&_
				"��<font color='#FF0000'>������Ϣ��</font>"&rs("Info")&"<BR>"
		rs.MoveNext
	Loop
Else
	Response.write "û�м�¼��"
End If
rs.close
Set rs=Nothing
Conn.Close
Set Conn=Nothing
%>
</body>
</html>