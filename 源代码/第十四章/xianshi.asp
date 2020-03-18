<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Connections/connection.asp" -->
<%
' IIf implementation
Function MM_IIf(condition, ifTrue, ifFalse)
  If condition = "" Then
    MM_IIf = ifFalse
  Else
    MM_IIf = ifTrue
  End If
End Function
%>

<%
Dim Recordset1__MMColParam
Recordset1__MMColParam = "119"
If (Request.QueryString("nid") <> "") Then 
  Recordset1__MMColParam = Request.QueryString("nid")
End If
%>
<%
Dim Recordset1
Dim Recordset1_cmd
Dim Recordset1_numRows

Set Recordset1_cmd = Server.CreateObject ("ADODB.Command")
Recordset1_cmd.ActiveConnection = MM_connection_STRING
Recordset1_cmd.CommandText = "SELECT * FROM new WHERE nid = ?" 
Recordset1_cmd.Prepared = true
Recordset1_cmd.Parameters.Append Recordset1_cmd.CreateParameter("param1", 5, 1, -1, Recordset1__MMColParam) ' adDouble

Set Recordset1 = Recordset1_cmd.Execute
Recordset1_numRows = 0
%>
<%

Set Command1 = Server.CreateObject ("ADODB.Command")
Command1.ActiveConnection = MM_connection_STRING
Command1.CommandText = "UPDATE new  SET hits = hits+1  WHERE nid ="&Recordset1__MMColParam
Command1.CommandType = 1
Command1.CommandTimeout = 0
Command1.Prepared = true
Command1.Execute()

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>首页</title>
<link href="css.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
.STYLE2 {font-size: 10pt}
.STYLE6 {font-size: 18px}
.STYLE8 {
	font-size: 14;
	color: #000000;
	font-family: "宋体";
}
-->
</style>
</head>
<body leftmargin="0" topmargin="0">
<!--#include file="head.asp" -->
<table width="770" border="0" align="center" cellpadding="1" bgcolor="#6687BA">
  <tr>
    <td height="400" valign="top" bgcolor="#FFFFFF"><p align="center" class="STYLE6"><br />
      <%=(Recordset1.Fields.Item("title").Value)%></p>
    <hr align="center" width="80%" />
    <p align="center">作者：<%=(Recordset1.Fields.Item("zuozhe").Value)%> &nbsp;&nbsp;<%=(Recordset1.Fields.Item("dateandtime").Value)%>点击次数:<%=(Recordset1.Fields.Item("hits").Value)%></p>
    <table width="80%" border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td>
		<% If(Recordset1.Fields.Item("selectpic").Value=true) Then  
	%>
		<% End If %>&nbsp;&nbsp;<span class="STYLE8"><%=(Recordset1.Fields.Item("content").Value)%></span><img src="<%=(Recordset1.Fields.Item("picurl").Value)%>" alt="" name="img" width="220" height="220" align="left" id="img" /></td>
      </tr>
    </table>
    </p></td>
  </tr>
</table>
<table width="770" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#6687BA">
  <tr>
    <td height="27" bgcolor="#FFFFFF"><div align="center" class="STYLE2">copyright&copy;new center </div></td>
  </tr>
</table>
</body>
</html>
<%
Recordset1.Close()
Set Recordset1 = Nothing
%>
