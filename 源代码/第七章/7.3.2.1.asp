<!--#include file="adovbs.inc"-->
<html>

<head>
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
'����ÿҳ��ʾ��¼����Ŀ
rs.PageSize=2
Dim page
page=rs.PageCount
PageNo=Trim(Request.QueryString("Page"))
If PageNo="" Then PageNo=1
PageNo=Cint(PageNo)
If PageNo<1 Then PageNo=1
If PageNo>Page Then PageNo=Page
rs.AbsolutePage=PageNo
For i=1 To rs.PageSize
	If rs.EOF then exit for
	Response.write "<font color='#FF0000'>ְλ��</font>"&rs("Name")&_
				"��<font color='#FF0000'>������Ϣ��</font>"&rs("Info")&"<BR>"
	rs.movenext
next
For i=1 to Page
	Response.write "<a href='13.3.2.asp?Page="&i&"'>��"&i&"ҳ</a>  "
Next
rs.close
Set rs=Nothing
Conn.Close
Set Conn=Nothing
%>
</body>

</html>
