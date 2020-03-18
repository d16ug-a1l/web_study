<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../Connections/connection.asp" -->
<%
dim success
success=request.Form("checkbox")
 If(success="")  Then success=false
Set Command1 = Server.CreateObject ("ADODB.Command")
Command1.ActiveConnection = MM_connection_STRING
Command1.CommandText = "UPDATE  yuding  SET yd_ifsuccess="&success&"  WHERE id="&request.Form("id")
Command1.Execute()
response.Write("<script language=javascript>history.go(-1)</script>")
%>

