<html>

<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>���ֶ�����</title>
</head>

<body>

<form method="POST" action="Search.asp">
	<p align="center"><font face="�����п�" size="6" color="#0000FF">�� �� ģ ��</font></p>
	<p align="center"><font face="�����п�" color="#0000FF">���ֶ�</font><select size="1" name="Field">
	<%
	  Session("prev_search")=""
	  Set Conn=Server.CreateObject("ADODB.Connection")
      'Response.Write(Server.MapPath("User.mdb")&"<BR>")
	  Conn.ConnectionString="Provider=Microsoft.Jet.OLEDB.4.0;"&_
  			"Data Source="&Server.MapPath("book.mdb")
	  Conn.Open
	  Sql="select * from Book_Info"
	  set rs=Conn.Execute(Sql) 
	  If rs.EOF=false Then  
         For i=0 To rs.Fields.Count-1
         	Response.write("<option value='"&rs.Fields(i).Name&","&rs.Fields(i).type&"'>"&rs.Fields(i).Name&"</option>")
         Next   	  	 
   	  End If
	%>
	</select><font face="�����п�" color="#0000FF">����</font>
	<input type="text" name="Content" size="16">
	<select size="1" name="Type">
	<option selected value="MH">ģ��</option>
	<option value="JQ">��ȷ</option>
	</select><input type="submit" value="����" name="B1"></p>
	<p align="center"><input type="radio" value="Reset" checked name="R1">��������<input type="radio" name="R1" value="Result">�ڽ���в�ѯ</p>
</form>
 
</body>

</html>