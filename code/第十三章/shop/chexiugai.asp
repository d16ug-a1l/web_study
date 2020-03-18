<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="Connections/connection.asp" -->
<%
set con=server.CreateObject("adodb.connection")
con.mode=2
con.open(MM_connection_STRING)
i=request.QueryString.Count
for j=1 to i step 1
num=trim(Request.QueryString.Item(j))
did=trim(Request.QueryString.Key(j)) 
Set rs=Server.CreateObject("ADODB.recordset")
sql= "select * from dingdan WHERE id= "&did
rs.open sql,con,1,3
rs("num")=num
rs.update
next
%>
<%
rs=nothing
rs.close()
Response.Redirect("che.asp")
%>
