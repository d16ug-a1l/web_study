<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../Connections/connection.asp" -->
<%
Set Command1 = Server.CreateObject ("ADODB.Command")
Command1.ActiveConnection = MM_connection_STRING
Command1.CommandText = "DELETE FROM fenlei  WHERE id ="&request.querystring("id")
Command1.Execute()
response.Redirect("fenlei.asp")
%>

