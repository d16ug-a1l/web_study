<%
ID=Trim(Request.QueryString("ID"))
If ID="" Then 
	Response.End
End If
ID=Cint(ID)
Typea=Request.QueryString("Type")
If Typea="" Then 
	Typea="LM"
End If
If Typea="LM" Then 
	Forum="Res_Info"
	ForumID="ResID"
	ForumName="Name"
ElseIf Typea="File" Then
	Forum="File_Info"
	ForumID="FileID"
	ForumName="Content"
End If
Content=Request.Form("T1")
If trim(Content)="" Then
	Response.Write("栏目名称或者文章标题不能为空！")
	Response.End
End If
SelectGroup=CInt(Request.Form("SelectGroup"))
SelectUser=Cint(Request.Form("SelectUser"))
'Response.write("SelectUser is "&SelectUser&"<BR>")
val=split(Request.Form("C"),",")
nTemp1=0
nTemp2=0
nTemp3=0
for each v in val 
	nval=Cint(trim(v))
	
	If nval>=100 Then
		If Int(nval/100)>nTemp1 Then nTemp1=Int(nval/100)
	ElseIf nval>=10 Then
		If Int(nval/10)>nTemp2 Then nTemp2=Int(nval/10)
	ElseIf nval>=1 Then
		If nval>nTemp3 Then nTemp3=nval
	End IF
Next
nval=nTemp1*100+nTemp2*10+nTemp3
If nval=0 Then
	nval=771
End If

val=split(Request.Form("U"),",")
nTemp1=0
for each v in val 
	nTval=Cint(trim(v))
	If nTval>nTemp1 Then nTemp1=nTval
Next
If nTemp1=0 Then nTemp1=1
'Response.write(nTemp1&"<BR>")
SelectAction=Request.Form("SelectAction")
Set Conn=Server.Createobject("Adodb.Connection") 
Conn.ConnectionString="Provider=Microsoft.Jet.OLEDB.4.0;"&_
 			"Data Source="&Server.MapPath("user.mdb")
Conn.Open
'Response.write(SelectAction&"<BR>")
If SelectAction="Add" Then
	Set rs=Conn.Execute("Select * from "&Forum&" where "&ForumName&"='"&Content&"'")
	'Response.write(Sql&"<BR>")
	If rs.EOF =false Then 
		ErrMsg="该项目已经存在！"
		Response.write(ErrMsg&"<BR>")

	Else 
		Sql="Insert into "&Forum&"("&ForumName&",Owner"
		'Response.write(Sql&"<BR>")
		If Typea="File" Then
			Sql=Sql&",LanMuID,Allow) values('"&Content&"',"&Session("Id")&","&Request.Form("LMID")&",'"&nVal&"')"	
			Conn.Execute(Sql)	
			Set rs=Conn.Execute("select ID from File_Info where Content='"&Content&"'")
			ID=CInt(rs("ID"))
		Else
			Sql=Sql&") values('"&Content&"',"&Session("Id")&")"
			Conn.Execute(Sql)
			Set rs=Conn.Execute("select ID from Res_Info where Name='"&Content&"'")
			ID=CInt(rs("ID"))
			Set rs=Conn.Execute("Select * from Group_Role where GroupID="&SelectUser&" and "&ForumID&"="&ID)
			If rs.EOF=false Then
				Sql="Update [Group_Role] Set [GroupID]="&SelectUser&",["&ForumID&"]="&ID&",[Action]='"&nTemp1&"'"
			Else	
				Sql="Insert into [Group_Role]([GroupID],[ResID],[Action]) values("&SelectGroup&","&ID&",'"&nVal&"')"
			End If
			Conn.Execute (Sql)	
		End If
		Set rs=Conn.Execute("Select * from User_Role where UserID="&SelectUser&" and "&ForumID&"="&ID)
		If rs.EOF=false Then
			Sql="Update [User_Role] Set [UserID]="&SelectUser&",["&ForumID&"]="&ID&",[Action]='"&nTemp1&"'"
		Else
			Sql="Insert into [User_Role]([UserID],["&ForumID&"],[Action]) values("&SelectUser&","&ID&",'"&nTemp1&"')"
		End If
		Conn.Execute (Sql)	
	End If
	
ElseIf SelectAction="Modify" Then
	Sql="Update ["&Forum&"] Set [" &ForumName&"]='"&Content&"',[Owner]="&Session("Id")
 	If Typea="File" Then
		Sql=Sql&",[LanMuID]="&Request.Form("LMID")&",[Allow]='"&nVal&"' where ID="&ID
		Conn.Execute(Sql)		
	Else
		Sql=Sql&" where ID="&ID
		Conn.Execute(Sql)	
		Sql="Update [Group_Role] set [GroupID]="&SelectGroup&",[ResID]="&ID&",[Action]="&nVal&" where GroupID="&SelectGroup&" and ResID="&ID
		Conn.Execute (Sql)	
	End If
	Sql="Update User_Role set [UserID]="&SelectUser&",["&ForumID&"]="&ID&",[Action]="&nTemp1&" where UserID="&SelectUser&" and "&ForumID&"="&ID
	Conn.Execute(Sql)	 

ElseIf SelectAction="Delete" Then
	 Sql="delete from "&Forum&" where ID="&ID
	 Conn.Execute(Sql)
	 Sql="delete from Group_Role where "&ForumID&"="&ID
	 Conn.Execute(Sql)
	 Sql="delete from User_Role where "&ForumID&"="&ID
	 Conn.Execute(Sql)
	 Response.write("删除成功")
 
Else
	Response.Write("动作选择错误")
 
End If
%>
