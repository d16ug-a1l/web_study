<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Connections/connection.asp" -->
<%
Dim MM_editAction
MM_editAction = CStr(Request.ServerVariables("SCRIPT_NAME"))
If (Request.QueryString <> "") Then
  MM_editAction = MM_editAction & "?" & Server.HTMLEncode(Request.QueryString)
End If

' boolean to abort record edit
Dim MM_abortEdit
MM_abortEdit = false
%>
<%
' *** Delete Record: construct a sql delete statement and execute it

If (CStr(Request("MM_delete")) = "form1" And CStr(Request("MM_recordId")) <> "") Then

  If (Not MM_abortEdit) Then
    ' execute the delete
    Set MM_editCmd = Server.CreateObject ("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_connection_STRING
    MM_editCmd.CommandText = "DELETE FROM postMain WHERE main_id = ?"
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 5, 1, -1, Request.Form("MM_recordId")) ' adDouble
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close

    ' append the query string to the redirect URL
    Dim MM_editRedirectUrl
    MM_editRedirectUrl = "admingl.asp"
    If (Request.QueryString <> "") Then
      If (InStr(1, MM_editRedirectUrl, "?", vbTextCompare) = 0) Then
        MM_editRedirectUrl = MM_editRedirectUrl & "?" & Request.QueryString
      Else
        MM_editRedirectUrl = MM_editRedirectUrl & "&" & Request.QueryString
      End If
    End If
    Response.Redirect(MM_editRedirectUrl)
  End If

End If
%>
<%
Dim Recordset1__MMColParam
Recordset1__MMColParam = "1"
If (Request.QueryString("id") <> "") Then 
  Recordset1__MMColParam = Request.QueryString("id")
End If
%>
<%
Dim Recordset1
Dim Recordset1_cmd
Dim Recordset1_numRows

Set Recordset1_cmd = Server.CreateObject ("ADODB.Command")
Recordset1_cmd.ActiveConnection = MM_connection_STRING
Recordset1_cmd.CommandText = "SELECT * FROM postMain WHERE main_id = ?" 
Recordset1_cmd.Prepared = true
Recordset1_cmd.Parameters.Append Recordset1_cmd.CreateParameter("param1", 5, 1, -1, Recordset1__MMColParam) ' adDouble

Set Recordset1 = Recordset1_cmd.Execute
Recordset1_numRows = 0
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>删除帐号</title>
<style type="text/css">
<!--
.zt {
	font-family: "宋体";
	font-size: 14px;
	color: #205E7B;
}
.wb {
	font-family: "宋体";
	font-size: 12px;
	color: #FFFFFF;
}
-->
</style>
</head>

<body>
<table width="95%" border="0" cellspacing="1" cellpadding="0" align="center">
  <tr>
    <td height="20" class="zt"><div align="center">您确定要删除帖子吗？</div></td>
  </tr>
</table>
<table width="95%" border="0" cellspacing="1" cellpadding="0" align="center">
  <tr>
    <td height="50" class="zt">
    <form id="form1" name="form1" method="POST" action="<%=MM_editAction%>">
      
        <div align="center">
          <input type="submit" name="button" id="button" value="确定" />
          <a href="javascript:history.go(-1)">返回</a></div>
    
        <input type="hidden" name="MM_delete" value="form1" />
        <input type="hidden" name="MM_recordId" value="<%= Recordset1.Fields.Item("main_id").Value %>" />
    </form>    </td>
  </tr>
</table>
<br />
</body>
</html>
<%
Recordset1.Close()
Set Recordset1 = Nothing
%>
