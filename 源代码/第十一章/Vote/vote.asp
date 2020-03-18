<%
 Set Conn=Server.Createobject("Adodb.Connection") 
 Conn.ConnectionString="Provider=Microsoft.Jet.OLEDB.4.0;"&_
  			"Data Source="&Server.MapPath("vote.mdb")
 Conn.Open
ip=Request("REMOTE_ADDR")

Sql="Select * From VoteIP where IP='"&ip&"'"
Set rs=Conn.Execute(Sql)
If rs.EOF=False Then
	Response.write("该IP地址已经投票过，请勿重复投票")
Else
	strVote=Split(Request.form("checkbox"),",") 
	Sql="Insert Into VoteIP(IP) Values('"&ip&"')"
	Conn.Execute(Sql)
	for each str in strVote 
	str=trim(str)
	Sql="Update [VoteItem] Set [Count]=[Count]+1 where IsDisp='1' and Item='"& str&"'" 
    Conn.Execute(Sql)
	next 
	Response.write("投票成功<BR>")
	Response.write("<a href='Result.asp'>查看投票结果</a>")
End IF
%>
