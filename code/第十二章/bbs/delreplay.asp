<%@LANGUAGE="VBSCRIPT"%>
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
    MM_editCmd.CommandText = "DELETE FROM postRe WHERE re_id = ?"
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 5, 1, -1, Request.Form("MM_recordId")) ' adDouble
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close

    ' append the query string to the redirect URL
    Dim MM_editRedirectUrl
    MM_editRedirectUrl = "adminglxg.asp?main_id=m_id"
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
Recordset1_cmd.CommandText = "SELECT * FROM postRe WHERE re_id = ?" 
Recordset1_cmd.Prepared = true
Recordset1_cmd.Parameters.Append Recordset1_cmd.CreateParameter("param1", 5, 1, -1, Recordset1__MMColParam) ' adDouble

Set Recordset1 = Recordset1_cmd.Execute
Recordset1_numRows = 0
%>
<style type="text/css">
<!--
.bg {
	background-repeat: no-repeat;
}
.STYLE1 {font-size: 12px}
.STYLE2 {font-size: 14px}
-->
</style>
<br />
<table width="98%" border="0" cellspacing="0" cellpadding="0" align="center"style="BORDER-LEFT: #000000 1px solid; BORDER-RIGHT: #000000 1px solid" bgcolor="#64B3D9" >
  <tr>
    <td height="50" valign="top"><form ACTION="<%=MM_editAction%>" METHOD="POST" id="form1" name="form1">
      <label>
       <div align="center"><br />
         你确定要删除这条回复吗？ 
           <input type="submit" name="button" id="button" value="确定" />
          <a href="javascript:history.go(-1)">返回</a></div>
      </label>
      <label>
      </label>
      <input type="hidden" name="MM_delete" value="form1" />
      <input type="hidden" name="MM_recordId" value="<%= Recordset1.Fields.Item("re_id").Value %>" />
</form>
    </td>
  </tr>
</table>
<table width="98%" border="0" cellspacing="0" cellpadding="0" align="center" bgcolor="#000000">
  <tr>
    <td align="center" height="30" class=td1>
    </td>
  </tr>
</table>
</body>
</html>
<%
Recordset1.Close()
Set Recordset1 = Nothing
%>
