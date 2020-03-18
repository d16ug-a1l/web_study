<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="Connections/connection.asp" -->
<%
Dim Recordset1
Dim Recordset1_cmd
Dim Recordset1_numRows

Set Recordset1_cmd = Server.CreateObject ("ADODB.Command")
Recordset1_cmd.ActiveConnection = MM_connection_STRING
Recordset1_cmd.CommandText = "SELECT * FROM orderid" 
Recordset1_cmd.Prepared = true

Set Recordset1 = Recordset1_cmd.Execute
Recordset1_numRows = 0
%>
<%

Set Command1 = Server.CreateObject ("ADODB.Command")
Command1.ActiveConnection = MM_connection_STRING
Command1.CommandText = "UPDATE orderid  SET orderid=orderid+1 "
Command1.CommandType = 1
Command1.CommandTimeout = 0
Command1.Prepared = true
Command1.Execute()
%>
<%
Session("orderid")=Recordset1.Fields.Item("orderid").Value
response.Redirect("addche.asp?id="&request.QueryString("id"))
%>
<%
Recordset1.Close()
Set Recordset1 = Nothing
%>
