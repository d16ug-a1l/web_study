<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
' *** Logout the current user.
MM_Logout = CStr(Request.ServerVariables("URL")) & "?MM_Logoutnow=1"
If (CStr(Request("MM_Logoutnow")) = "1") Then
  Session.Contents.Remove("MM_Username")
  Session.Contents.Remove("MM_UserAuthorization")
  MM_logoutRedirectPage = "login.asp"
  ' redirect with URL parameters (remove the "MM_Logoutnow" query param).
  if (MM_logoutRedirectPage = "") Then MM_logoutRedirectPage = CStr(Request.ServerVariables("URL"))
  If (InStr(1, UC_redirectPage, "?", vbTextCompare) = 0 And Request.QueryString <> "") Then
    MM_newQS = "?"
    For Each Item In Request.QueryString
      If (Item <> "MM_Logoutnow") Then
        If (Len(MM_newQS) > 1) Then MM_newQS = MM_newQS & "&"
        MM_newQS = MM_newQS & Item & "=" & Server.URLencode(Request.QueryString(Item))
      End If
    Next
    if (Len(MM_newQS) > 1) Then MM_logoutRedirectPage = MM_logoutRedirectPage & MM_newQS
  End If
  Response.Redirect(MM_logoutRedirectPage)
End If
%>
<!--#include file="Connections/connection.asp" -->
<%
if(Session("MM_Username")="")then
response.Redirect("login.asp")
end if
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
If (CStr(Request("MM_update")) = "form1") Then
  If (Not MM_abortEdit) Then
    ' execute the update
    Dim MM_editCmd

    Set MM_editCmd = Server.CreateObject ("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_connection_STRING
    MM_editCmd.CommandText = "UPDATE passadmin SET usrname = ?, passwd = ? WHERE usrname = ?" 
    MM_editCmd.Prepared = true
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 202, 1, 50, Request.Form("username")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param2", 202, 1, 50, Request.Form("pass")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param3", 200, 1, 50, Request.Form("MM_recordId")) ' adVarChar
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close

    ' append the query string to the redirect URL
    Dim MM_editRedirectUrl
    MM_editRedirectUrl = "adminuser.asp"
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
If (Request.QueryString("username") <> "") Then 
  Recordset1__MMColParam = Request.QueryString("username")
End If
%>
<%
Dim Recordset1
Dim Recordset1_cmd
Dim Recordset1_numRows

Set Recordset1_cmd = Server.CreateObject ("ADODB.Command")
Recordset1_cmd.ActiveConnection = MM_connection_STRING
Recordset1_cmd.CommandText = "SELECT * FROM passadmin WHERE usrname = ?" 
Recordset1_cmd.Prepared = true
Recordset1_cmd.Parameters.Append Recordset1_cmd.CreateParameter("param1", 200, 1, 50, Recordset1__MMColParam) ' adVarChar

Set Recordset1 = Recordset1_cmd.Execute
Recordset1_numRows = 0
%><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>无标题文档</title>
<style type="text/css">
<!--
.zt {
	font-family: "宋体";
	font-size: 14px;
	color: #205E7B;
}
.STYLE1 {color: #FFFFFF}
-->
</style>
</head>

<body>
<table width="95%" border="0" cellspacing="1" cellpadding="0" align="center">
  <tr>
    <td height="50" class="zt"><a href="<%= MM_Logout %>">退出</a>&nbsp;&nbsp;<a href="adminuser.asp">管理员帐号管理</a>&nbsp;&nbsp;<a href="admingl.asp">帖子管理</a>&nbsp;&nbsp;</td>
  </tr>
</table>
<br />
<table width="96%" border="0" align="center" cellspacing="1" bgcolor="#205E7B">
  <tr align="center">
    <td  height="25"><div align="left" class="STYLE1">修改帐号和密码</div></td>
  </tr>
  <tr>
    <td height="1" colspan="3" bgcolor="#000000"></td>
  </tr>
<tr>
          <td height="35" colspan="7" bgcolor="#D6DFF7"><form ACTION="<%=MM_editAction%>"  method="POST" name="form1">
            <table width="96%" border="0" cellspacing="1" cellpadding="3" align="center">
              <tr>
                <td>帐号名称：
                  <label>
                  <input name="username" type="text" value=<%= Request.QueryString("username")%> />
                </label></td>
              </tr>
              <tr>
                <td>修改密码：
                  <input type="text" name="pass" value="" /></td>
              </tr>
              <tr>
                <td><input type="submit" name="admin" value="确认" onclick="check()" />
                &nbsp;
                <label>
                <input type="reset" name="button" id="button" value="重置" />
                &nbsp;&nbsp;<a href="javascript:history.go(-1)">返回</a></label></td>
              </tr>
            </table>
          
          
            <input type="hidden" name="MM_update" value="form1" />
            <input type="hidden" name="MM_recordId" value="<%=Recordset1.Fields.Item("usrname").Value %>" />
</form></td>
  </tr>
</table>
<script language="javascript">
function check(){
if(document.form1.pass.value==""||document.form1.pass1.value=="")
{alert("请输入密码")}
else{
if(document.form1.pass.value!=document.form1.pass1.value)
{
alert("两次密码输入不正确")
}
}
}

</script>
</body>
</html>
<%
Recordset1.Close()
Set Recordset1 = Nothing
%>
