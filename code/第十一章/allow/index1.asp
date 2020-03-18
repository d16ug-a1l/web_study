<%
GetResAllow("¹«¸æÀ¸","2","2")

 
Public function GetResAllow(strMessage,UserID,GroupID)

Set Conn=Server.Createobject("Adodb.Connection") 
Conn.ConnectionString="Provider=Microsoft.Jet.OLEDB.4.0;"&_
 			"Data Source="&Server.MapPath("user.mdb")
Conn.Open
strMessage=trim(strMessage)
Response.write("strMessage  is "&strMessage&"<BR>")
Sql="Select * From Res_Info where Name='"&strMessage&"'"
Set rs=Conn.Execute(Sql)
If rs.EOF=False Then
	OwnerID=rs("Owner")
	ResID=rs("ID")
Else
	Response.End
End If
Response.write("OwnerIDis "&OwnerID&"<BR>")
Dim OwnerGroup
Sql="Select Group From User_Info where ID='"&OwnerID&"'"
Set rs=Conn.Execute(Sql)
If rs.EOF=False Then
	OwnerGroup=rs("Group")
Else
	Response.End
End If
Response.write("OwnerGroup is "&OwnerGroup&"<BR>")
Dim GroupAction,UserAction
Sql="Select Action From Group_Role where Name='"&GroupID&"' ResID='"&ResID&"'"
Set rs=Conn.Execute(Sql)
If rs.EOF=False Then
	GroupAction=rs("Action")
End If 
Response.write("GroupAction is "&GroupAction&"<BR>")
Sql="Select Action From User_Role where Name='"&UserID&"' ResID='"&ResID&"'"
Set rs=Conn.Execute(Sql)
If rs.EOF=False Then
	UserAction=rs("Action")
Else
	UserAction=""
End If
Response.write("UserAction is "&UserAction&"<BR>")
Dim nUser,nGroup,Result
If UserID=OwnerID Then
	If UserAction="" Then
		Result=Cint(Mid(GroupAction,1,1)
	Else
		Result=Cint(Mid(UserAction,1,1))
	End If
ElseIf GroupID=OwnerGroup Then
	If UserAction="" Then
		Result=Cint(Mid(GroupAction,2,1)
	Else
		Result=Cint(Mid(UserAction,2,1))
	End If
Else
	If UserAction="" Then
		Result=Cint(Mid(GroupAction,3,1)
	Else
		Result=Cint(Mid(UserAction,3,1))
	End If
End If
GetResAllow=Result
End function
%>

