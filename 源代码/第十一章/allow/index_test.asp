<%
Set Conn=Server.Createobject("Adodb.Connection") 
Conn.ConnectionString="Provider=Microsoft.Jet.OLEDB.4.0;"&_
 			"Data Source="&Server.MapPath("user.mdb")
Conn.Open
Sql="Update [File_Info] set [Content]='��Һã�Asp��ȫ������122112',[Owner]=1,[LanMuID]=1,[Allow]='771' where ID=2"
Conn.Execute(Sql)

 
%>