<!--#include file="funciton.asp"-->
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�½���ҳ 1</title>
<base target="_self">
</head>

<body bgcolor="#99CCFF">
<p align="center"><font face="�����п�" size="6" color="#0000FF">Ȩ �� �� �� ģ ��</font></p> 
<%
ID=Request.QueryString("ID")
If ID="" Then 
	Response.End
End If
TypeRes=Request.QueryString("Type")
If TypeRes="" Then 
	TypeRes="LM"
End If
GroupID=CInt(Session("GroupID"))

Sql="Select * From "
If TypeRes="LM" Then 
	Sql=Sql& " Res_Info where ID="&ID
ElseIf TypeRes="File" Then
	Sql=Sql& " File_Info where ID="&ID
End If
Set Conn=Server.Createobject("Adodb.Connection") 
Conn.ConnectionString="Provider=Microsoft.Jet.OLEDB.4.0;"&_
 			"Data Source="&Server.MapPath("user.mdb")
Conn.Open

Set rs=Conn.Execute(Sql)
If rs.EOF = False Then
	If TypeRes="LM" Then
		Response.write("<P align='center'><font face='�����п�' size='6' color='#0000FF'>"&rs("Name")&"</font></p> ")
		Dim strName
		strName=rs("Name")
		nCode=Cint(GetResAllow(strName,Session("Id"),GroupID))	
		If nCode>1 Then
			Response.write("<form method='POST' action='modify.asp?Type=LM&ID="&ID&"'><select size='1' name='SelectAction'>")
			Response.write("<option value='Add'>����</option>")
			Response.write("<option value='Modify'>�޸�</option>")	
			If nCode=7 Then
			   Response.write("<option value='Delete'>ɾ��</option>")
			End If
			Response.write(" </select> ")
			Response.write(" <p>��Ŀ���ƣ�<input type='LanMu' name='T1' size='20' value='"& strName &"'></p>")
%>

<p>���û���<select size="1" name="SelectGroup">
<%
			Sql="Select * From Group_Info where ID>"&GroupID
			Set rs=Conn.Execute(Sql)
			Do while rs.EOF=False
				Response.write("<option value='"&rs("ID")&"'>"&rs("Name")&"</option>")
				rs.movenext			
			loop
			
%>
</select>
<p>��Ȩ�ޣ�&nbsp; 
<input type="checkbox" name="C" value="100" >�ļ����ۿ�
<input type="checkbox" name="C" value="300" >�ļ����޸�
<input type="checkbox" name="C" value="700" >�ļ�����ɾ��</p>
<p>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<input type="checkbox" name="C" value="10" >ͬ���û��ۿ�
<input type="checkbox" name="C" value="30" >ͬ���û��޸�
<input type="checkbox" name="C" value="70" >ͬ���û�ɾ��</p>
<p>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<input type="checkbox" name="C" value="1" >�����û��ۿ�
<input type="checkbox" name="C" value="3" >�����û��޸�
<input type="checkbox" name="C" value="7" >�����û�ɾ��</p>
<p>���������û��Ա���ĿȨ��</p>
<p>�û���<select size="1" name="SelectUser">
<%
			Sql="Select * From User_Info where GroupID>"&GroupID
			Set rs=Conn.Execute(Sql)
			Do while rs.EOF=False
				Response.write("<option value='"&rs("ID")&"'>"&rs("user")&"</option>")
				rs.movenext			
			loop
			
%>
</select><input type="checkbox" name="U" value="1">�ۿ�
<input type="checkbox" name="U" value="3">�޸�
<input type="checkbox" name="U" value="7">ɾ��</p>
<p><input type="submit" value="ȷ��" name="B1"><input type="reset" value="����" name="B2"></p>	
</form>
<%
		End If 
	End If 
	If TypeRes="File" Then 
		strName=rs("Content")
		Allow=Request.QueryString("Allow")
		OwnerID=Request.QueryString("Owner")
		Response.write("<P align='center'><font face='�����п�' size='6' color='#0000FF'>�����ļ�Ȩ��</font></p> ")
		nCode=Cint(GetFileAllow(ID, CInt(Session("Id")),GroupID,OwnerID,Allow))
		If nCode>1 Then
			Response.write("<form method='POST' action='modify.asp?Type=File&ID="&ID&"'><select size='1' name='SelectAction'>")
			Response.write("<option value='Add'>����</option>")
			Response.write("<option value='Modify'>�޸�</option>")	
			If nCode=7 Then
				Response.write("<option value='Delete'>ɾ��</option>")
			End If
			Response.write(" </select> ")
%>
<p>�ļ����ݣ�<input type="text" name="T1" size="50" value=<% Response.write(strName) %>></p>
<p>������Ŀ��<select size="1" name="LMID">
<%
			Sql="Select * From Res_Info  "
			Set rs=Conn.Execute(Sql)
			Do while rs.EOF=False
				Response.write("<option value='"&rs("ID")&"'>"&rs("Name")&"</option>")
				rs.movenext			
			loop
			
%>
</select>
<p>��Ȩ�ޣ�&nbsp; 
<input type="checkbox" name="C" value="100" >�ļ����ۿ�
<input type="checkbox" name="C" value="300" >�ļ����޸�
<input type="checkbox" name="C" value="700" >�ļ�����ɾ��</p>
<p>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<input type="checkbox" name="C" value="10" >ͬ���û��ۿ�
<input type="checkbox" name="C" value="30" >ͬ���û��޸�
<input type="checkbox" name="C" value="70" >ͬ���û�ɾ��</p>
<p>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<input type="checkbox" name="C" value="1" >�����û��ۿ�
<input type="checkbox" name="C" value="3" >�����û��޸�
<input type="checkbox" name="C" value="7" >�����û�ɾ��</p>
<p>�û���<select size="1" name="SelectUser">

<%
			Sql="Select * From User_Info where GroupID>"&GroupID
			Set rs=Conn.Execute(Sql)
			Do while rs.EOF=False
				Response.write("<option value='"&rs("ID")&"'>"&rs("user")&"</option>")
				rs.movenext			
			loop
%>
</select><input type="checkbox" name="U" value="1">�ۿ�<input type="checkbox" name="U" value="3">�޸�<input type="checkbox" name="U" value="7">ɾ��</p>
<p><input type="submit" value="ȷ��" name="B1"><input type="reset" value="����" name="B2"></p>	
</form>
<%
		End If
	End If
End If

%>

</body>

</html>