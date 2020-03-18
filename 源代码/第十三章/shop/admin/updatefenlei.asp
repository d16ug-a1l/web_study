<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../Connections/connection.asp" -->
<%
Set Command1 = Server.CreateObject ("ADODB.Command")
Command1.ActiveConnection = MM_connection_STRING
Command1.CommandText = "UPDATE fenlei  SET fenlei='"&request.Form("fenlei")&"' WHERE id ="&request.Form("fid")
Command1.Execute()
Command1.CommandText = "UPDATE shop  SET sp_leibie='"&request.Form("fenlei")&"' WHERE sp_leibie='"&request.Form("oldfenlei")&"'"
Command1.Execute()
response.Redirect("fenlei.asp")
%>


