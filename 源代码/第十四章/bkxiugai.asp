<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Connections/connection.asp" -->
<%
dim type1
type1=request.Form("type")
dim typeid
typeid=request.Form("typeid")
Set Command1 = Server.CreateObject ("ADODB.Command")
Command1.ActiveConnection = MM_connection_STRING
Command1.CommandText = "UPDATE bankuai  SET type ='"&type1&"' WHERE typeid="&typeid
Command1.CommandType = 1
Command1.CommandTimeout = 0
Command1.Prepared = true
Command1.Execute()
response.Redirect("bankuai.asp")
%>

