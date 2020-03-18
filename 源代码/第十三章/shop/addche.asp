<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="Connections/connection.asp" -->
<%
sp_id=Request.QueryString("id")
if(Session("orderid")="")then
response.Redirect("getorder.asp?id="&sp_id)
else
Set Command1 = Server.CreateObject ("ADODB.Command")
Command1.ActiveConnection = MM_connection_STRING
Command1.CommandText = "INSERT INTO dingdan (orderid, sp_id)  VALUES ('"&Session("orderid")&"','"&sp_id&"') "
Command1.CommandType = 1
Command1.CommandTimeout = 0
Command1.Prepared = true
Command1.Execute()
response.Write("<script language='javascript'>window.history.go(-1)</script>")
end if
%>
