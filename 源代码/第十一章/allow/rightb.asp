<!--#include file="funciton.asp"-->
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�½���ҳ 1</title>
</head>

<body>
<p align="center"><font face="�����п�" size="6" color="#0000FF">Ȩ �� �� �� ģ ��</font></p> 
<%
ID=Request.QueryString("ID")
If ID="" Then 
	Response.End
End If
Typea=Request.QueryString("Type")
If Typea="" Then 
	Typea="LM"
End If

Allow=Request.QueryString("Allow")
OwnerID=Request.QueryString("Owner")
Sql="Select * From "
If Typea="LM" Then 
	Sql=Sql& " Res_Info where ID="&ID
ElseIf Typea="File" Then
	Sql=Sql& " File_Info where ID="&ID
End If
Set Conn=Server.Createobject("Adodb.Connection") 
Conn.ConnectionString="Provider=Microsoft.Jet.OLEDB.4.0;"&_
 			"Data Source="&Server.MapPath("user.mdb")
Conn.Open
Response.write(Sql)
Set rs=Conn.Execute(Sql)
If rs.EOF = False Then
	If Typea="LM" Then 
		Response.write("align='center'><font face='�����п�' size='6' color='#0000FF'>"&rs("Name")&"</font></p> ")
		nCode=Cint(GetResAllow(rs("Name"),UserID,GroupID))
		If nCode>1 Then
			Response.write("<form method='POST' action=''><select size='1' name='D1'>")
			If nCode=3 Then
				Response.write("<option value=Add>����</option>")
				Response.write("<option value=Add>�޸�</option>")	
			ElseIf nCode=7 Then
				Response.write("<option value=Add>ɾ��</option>")
			End If
			Response.write(" </select> ")

		End If
	End If
End If

%>
<p><input type="submit" value="ȷ��" name="B1"><input type="reset" value="����" name="B2"></p>	
</form>
</body>

</html>