<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
' *** Logout the current user.
MM_Logout = CStr(Request.ServerVariables("URL")) & "?MM_Logoutnow=1"
If (CStr(Request("MM_Logoutnow")) = "1") Then
  Session.Contents.Remove("MM_Username")
  Session.Contents.Remove("MM_UserAuthorization")
  MM_logoutRedirectPage = "index.asp"
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
If (CStr(Request("MM_insert")) = "form1") Then
  If (Not MM_abortEdit) Then
    ' execute the insert
    Dim MM_editCmd

    Set MM_editCmd = Server.CreateObject ("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_connection_STRING
    MM_editCmd.CommandText = "INSERT INTO passadmin (usrname, passwd) VALUES (?, ?)" 
    MM_editCmd.Prepared = true
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 202, 1, 50, Request.Form("username")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param2", 202, 1, 50, Request.Form("textfield")) ' adVarWChar
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
if(Session("MM_Username")="")then
response.Redirect("login.asp")
end if
Dim Recordset1
Dim Recordset1_cmd
Dim Recordset1_numRows

Set Recordset1_cmd = Server.CreateObject ("ADODB.Command")
Recordset1_cmd.ActiveConnection = MM_connection_STRING
Recordset1_cmd.CommandText = "SELECT * FROM passadmin" 
Recordset1_cmd.Prepared = true

Set Recordset1 = Recordset1_cmd.Execute
Recordset1_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = -1
Repeat1__index = 0
Recordset1_numRows = Recordset1_numRows + Repeat1__numRows
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
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
    <td height="50" class="zt"><a href="<%= MM_Logout %>">退出</a>&nbsp;&nbsp;<a href="adminuser.asp">管理员帐号管理</a>&nbsp;&nbsp;<a href="admingl.asp">帖子管理</a>&nbsp;&nbsp;</td>
  </tr>
</table>
<table width="95%" border="0" cellspacing="1" cellpadding="0" align="center">
  <tr>
    <td height="50" class="zt"><form id="form1" name="form1" method="POST" action="<%=MM_editAction%>">
      用户名:
      <label>
      <input type="text" name="username" id="username" />
      </label>
      密码:
  <label>
  <input type="text" name="textfield" id="textfield" />
  </label>
  <label>
  <input type="submit" name="button" id="button" value="添加" />
  </label>
  &nbsp;
  <label>
  <input type="reset" name="button2" id="button2" value="重置" />
  </label>
  <input type="hidden" name="MM_insert" value="form1" />
</form></td>
  </tr>
</table>
<br />
<table width="96%" border="0" align="center" cellspacing="1" bgcolor="#205E7B">
  <tr align="center">
    <td width="272"  height="25"><font color="#FFFFFF" class="wb">管理员帐号（单击修改密码）</font></td>
    <td  height="25"><font color="#FFFFFF" class="wb">密码</font></td>
    <td height="25" colspan="3"><font color="#FFFFFF" class="wb">操作</font></td>
  </tr>
  <tr>
    <td height="1" colspan="7" bgcolor="#000000"></td>
  </tr>
  <% 
While ((Repeat1__numRows <> 0) AND (NOT Recordset1.EOF)) 
%>
    <tr bgcolor="#f0f0f0">
      <td height="25" align="center">      <label>
        <div align="center"><%=(Recordset1.Fields.Item("usrname").Value)%></div>        </label></td>
      <td width="280" height="25"><div align="center"><%=(Recordset1.Fields.Item("passwd").Value)%></div></td>
      <td width="143" height="25" align="center"><label><a href="adminuserxg.asp?username=<%=(Recordset1.Fields.Item("usrname").Value)%>">修改
      </a></label></td>
      <td width="149" height="25" align="center"><font color="#FF0000"><a href="adminuserdel.asp?username=<%=(Recordset1.Fields.Item("usrname").Value)%>">删除</a></font></td>
    </tr>
    <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  Recordset1.MoveNext()
Wend
%>

  <tr>
          <td colspan="11" bgcolor="#D6DFF7">&nbsp;</td>
  </tr>
</table>
</body>
</html>
<%
Recordset1.Close()
Set Recordset1 = Nothing
%>
