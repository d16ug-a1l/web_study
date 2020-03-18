<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Connections/connection.asp" -->
<%

Set Command1 = Server.CreateObject ("ADODB.Command")
Command1.ActiveConnection = MM_connection_STRING
Command1.CommandText = "DELETE FROM bankuai  WHERE typeid ="&request.querystring("typeid")
Command1.CommandType = 1
Command1.CommandTimeout = 0
Command1.Prepared = true
Command1.Execute()
response.Redirect("bankuai.asp")

%>

