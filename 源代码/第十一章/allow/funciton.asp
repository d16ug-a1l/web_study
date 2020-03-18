<%
Public Function OutPutFileContent(strMessage,UserID,GroupID)
Set Conn=Server.Createobject("Adodb.Connection") 
Conn.ConnectionString="Provider=Microsoft.Jet.OLEDB.4.0;"&_
 			"Data Source="&Server.MapPath("user.mdb")
Conn.Open
strMessage=trim(strMessage)
Sql="Select * From Res_Info where Name='"&strMessage&"'"
Set rs=Conn.Execute(Sql)
If rs.EOF=False Then
	OwnerID=rs("Owner")
Else
	Response.End
End If
Sql="Select * From File_Info where LanMuID="&CInt(rs("ID"))
Dim nCount
nCount=0
Set rs=Conn.Execute(Sql)
Do while rs.EOF=False  
  nCode=GetFileAllow(rs("ID"),UserID,GroupID,rs("Owner"),rs("Allow"))
  If nCode>=1 Then
    nCount=nCount+1
  	Response.write("<tr><td>")
  	Response.Write("."&rs("Content"))
  	Response.Write("</td></tr>")
  End If
  rs.movenext
loop
If nCount=0 Then
    Response.write("<tr><td>")
  	Response.Write("暂无该栏目信息" )
  	Response.Write("</td></tr>")
End If
OutPutFileContent=1
End Function






Public Function GetFileAllow(FileID,UserID,GroupID,OwnerID,Allow)
Set Conn1=Server.Createobject("Adodb.Connection") 
Conn1.ConnectionString="Provider=Microsoft.Jet.OLEDB.4.0;"&_
 			"Data Source="&Server.MapPath("user.mdb")
Conn1.Open

Dim GroupOfOwner
Sql="Select GroupID From User_Info where ID="&OwnerID
Set rs=Conn1.Execute(Sql)
If rs.EOF=False Then
	GroupOfOwner=rs("GroupID")
Else
	Response.End
End If


Dim nUser,nGroup,Result
Sql="Select * From Group_Info where ID="&CInt(GroupOfOwner)&" and GroupOwner="&CInt(UserID)
Set rs=Conn1.Execute(Sql)
If rs.EOF=False Then
	Result=7
Else
	Sql="Select Action From User_Role where UserID="&CInt(UserID)&" and FileID="&Cint(FileID)
	Set rs=Conn1.Execute(Sql)
	If rs.EOF=False Then
		UserAction=rs("Action")
	Else
		UserAction=""
	End If
	If UserID=OwnerID Then
		If UserAction="" Then
			Result=Cint(Mid(Allow,1,1))
		Else
			Result=Cint(Mid(UserAction,1,1))
		End If
	ElseIf GroupID=GroupOfOwner Then
		If UserAction="" Then
			Result=Cint(Mid(Allow,2,1))
		Else
			Result=Cint(Mid(UserAction,1,1))
		End If
	Else
		If UserAction="" Then
			Result=Cint(Mid(Allow,3,1))
		Else
			Result=Cint(Mid(UserAction,1,1))
		End If
	End If
End If
Conn1.close
GetFileAllow=Result
End Function





Public Function GetResAllow(strMessage,UserID,GroupID)
Set Conn=Server.Createobject("Adodb.Connection") 
Conn.ConnectionString="Provider=Microsoft.Jet.OLEDB.4.0;"&_
 			"Data Source="&Server.MapPath("user.mdb")
Conn.Open
strMessage=trim(strMessage)
Sql="Select * From Res_Info where Name='"&strMessage&"'"
Set rs=Conn.Execute(Sql)
If rs.EOF=False Then
	OwnerID=rs("Owner")
	ResID=rs("ID")
Else
	Response.End
End If
Dim Group_Owner
Sql="Select GroupID From User_Info where ID="&CInt(OwnerID)
Set rs=Conn.Execute(Sql)
If rs.EOF=False Then
	GroupOfOwner=rs("GroupID")
Else
	Response.End
End If
Sql="Select* From Group_Info where ID="&CInt(GroupOfOwner)&" and GroupOwner="&CInt(UserID)
Set rs=Conn.Execute(Sql)
If rs.EOF=False Then
	GetResAllow=7
End If
Dim GroupAction,UserAction
Sql="Select Action From Group_Role where GroupID="&CInt(GroupID)&" and ResID="&CInt(ResID)

Set rs=Conn.Execute(Sql)
If rs.EOF=False Then
	GroupAction=rs("Action")
End If 

Sql="Select Action From User_Role where UserID="&CInt(UserID)&" and ResID="&Cint(ResID)

Set rs=Conn.Execute(Sql)
If rs.EOF=False Then
	UserAction=rs("Action")
Else
	UserAction=""
End If

Dim nUser,nGroup,Result

If UserID=OwnerID Then
	If UserAction="" Then
		Result=Cint(Mid(GroupAction,1,1))
	Else
		Result=Cint(Mid(UserAction,1,1))
	End If
ElseIf GroupID=GroupOfOwner Then
	If UserAction="" Then
		Result=Cint(Mid(GroupAction,2,1))
	Else
		Result=Cint(Mid(UserAction,1,1))
	End If
Else
	If UserAction="" Then
		Result=Cint(Mid(GroupAction,3,1))
	Else
		Result=Cint(Mid(UserAction,1,1))
	End If
End If
'Conn.close
GetResAllow=Result

End Function
%>
