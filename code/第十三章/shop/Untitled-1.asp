<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="Connections/connection.asp" -->
<%
sp_id=Request.QueryString("id")
if(Session("orderid")="")then
response.Write("您还没有购物")
else
orderid=Session("orderid")
For Each objItem In Request.Form("num")
    num=Request.Form(Cstr(objItem)) 
    sid=Cstr(objItem) 
Set Command1 = Server.CreateObject ("ADODB.Command")
Command1.ActiveConnection = MM_connection_STRING
Command1.CommandText = "UPDATE dingdan  SET sp_num='"&num&"'  WHERE id='"&sid&"' "
Command1.CommandType = 1
Command1.CommandTimeout = 0
Command1.Prepared = true
Command1.Execute()
next
response.Redirect("che.asp")
end if
%>

