<!--#include file="funciton.asp"-->
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>新建网页 1</title>
<base target="_self">
</head>

<body bgcolor="#99CCFF">
<p align="center"><font face="华文行楷" size="6" color="#0000FF">权 限 管 理 模 块</font></p> 
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
		Response.write("<P align='center'><font face='华文行楷' size='6' color='#0000FF'>"&rs("Name")&"</font></p> ")
		Dim strName
		strName=rs("Name")
		nCode=Cint(GetResAllow(strName,Session("Id"),GroupID))	
		If nCode>1 Then
			Response.write("<form method='POST' action='modify.asp?Type=LM&ID="&ID&"'><select size='1' name='SelectAction'>")
			Response.write("<option value='Add'>增加</option>")
			Response.write("<option value='Modify'>修改</option>")	
			If nCode=7 Then
			   Response.write("<option value='Delete'>删除</option>")
			End If
			Response.write(" </select> ")
			Response.write(" <p>栏目名称：<input type='LanMu' name='T1' size='20' value='"& strName &"'></p>")
%>

<p>组用户：<select size="1" name="SelectGroup">
<%
			Sql="Select * From Group_Info where ID>"&GroupID
			Set rs=Conn.Execute(Sql)
			Do while rs.EOF=False
				Response.write("<option value='"&rs("ID")&"'>"&rs("Name")&"</option>")
				rs.movenext			
			loop
			
%>
</select>
<p>组权限：&nbsp; 
<input type="checkbox" name="C" value="100" >文件主观看
<input type="checkbox" name="C" value="300" >文件主修改
<input type="checkbox" name="C" value="700" >文件主户删除</p>
<p>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<input type="checkbox" name="C" value="10" >同组用户观看
<input type="checkbox" name="C" value="30" >同组用户修改
<input type="checkbox" name="C" value="70" >同组用户删除</p>
<p>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<input type="checkbox" name="C" value="1" >其他用户观看
<input type="checkbox" name="C" value="3" >其他用户修改
<input type="checkbox" name="C" value="7" >其他用户删除</p>
<p>设置其他用户对本栏目权限</p>
<p>用户：<select size="1" name="SelectUser">
<%
			Sql="Select * From User_Info where GroupID>"&GroupID
			Set rs=Conn.Execute(Sql)
			Do while rs.EOF=False
				Response.write("<option value='"&rs("ID")&"'>"&rs("user")&"</option>")
				rs.movenext			
			loop
			
%>
</select><input type="checkbox" name="U" value="1">观看
<input type="checkbox" name="U" value="3">修改
<input type="checkbox" name="U" value="7">删除</p>
<p><input type="submit" value="确定" name="B1"><input type="reset" value="重置" name="B2"></p>	
</form>
<%
		End If 
	End If 
	If TypeRes="File" Then 
		strName=rs("Content")
		Allow=Request.QueryString("Allow")
		OwnerID=Request.QueryString("Owner")
		Response.write("<P align='center'><font face='华文行楷' size='6' color='#0000FF'>设置文件权限</font></p> ")
		nCode=Cint(GetFileAllow(ID, CInt(Session("Id")),GroupID,OwnerID,Allow))
		If nCode>1 Then
			Response.write("<form method='POST' action='modify.asp?Type=File&ID="&ID&"'><select size='1' name='SelectAction'>")
			Response.write("<option value='Add'>增加</option>")
			Response.write("<option value='Modify'>修改</option>")	
			If nCode=7 Then
				Response.write("<option value='Delete'>删除</option>")
			End If
			Response.write(" </select> ")
%>
<p>文件内容：<input type="text" name="T1" size="50" value=<% Response.write(strName) %>></p>
<p>所属栏目：<select size="1" name="LMID">
<%
			Sql="Select * From Res_Info  "
			Set rs=Conn.Execute(Sql)
			Do while rs.EOF=False
				Response.write("<option value='"&rs("ID")&"'>"&rs("Name")&"</option>")
				rs.movenext			
			loop
			
%>
</select>
<p>组权限：&nbsp; 
<input type="checkbox" name="C" value="100" >文件主观看
<input type="checkbox" name="C" value="300" >文件主修改
<input type="checkbox" name="C" value="700" >文件主户删除</p>
<p>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<input type="checkbox" name="C" value="10" >同组用户观看
<input type="checkbox" name="C" value="30" >同组用户修改
<input type="checkbox" name="C" value="70" >同组用户删除</p>
<p>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<input type="checkbox" name="C" value="1" >其他用户观看
<input type="checkbox" name="C" value="3" >其他用户修改
<input type="checkbox" name="C" value="7" >其他用户删除</p>
<p>用户：<select size="1" name="SelectUser">

<%
			Sql="Select * From User_Info where GroupID>"&GroupID
			Set rs=Conn.Execute(Sql)
			Do while rs.EOF=False
				Response.write("<option value='"&rs("ID")&"'>"&rs("user")&"</option>")
				rs.movenext			
			loop
%>
</select><input type="checkbox" name="U" value="1">观看<input type="checkbox" name="U" value="3">修改<input type="checkbox" name="U" value="7">删除</p>
<p><input type="submit" value="确定" name="B1"><input type="reset" value="重置" name="B2"></p>	
</form>
<%
		End If
	End If
End If

%>

</body>

</html>