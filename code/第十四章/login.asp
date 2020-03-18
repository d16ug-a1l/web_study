<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Connections/connection.asp" -->
<%
' *** Validate request to log in to this site.
MM_LoginAction = Request.ServerVariables("URL")
If Request.QueryString <> "" Then MM_LoginAction = MM_LoginAction + "?" + Server.HTMLEncode(Request.QueryString)
MM_valUsername = CStr(Request.Form("UserName"))
If MM_valUsername <> "" Then
  Dim MM_fldUserAuthorization
  Dim MM_redirectLoginSuccess
  Dim MM_redirectLoginFailed
  Dim MM_loginSQL
  Dim MM_rsUser
  Dim MM_rsUser_cmd
  
  MM_fldUserAuthorization = ""
  MM_redirectLoginSuccess = "bankuai.asp"
  MM_redirectLoginFailed = "login.asp"

  MM_loginSQL = "SELECT Username, Passwd"
  If MM_fldUserAuthorization <> "" Then MM_loginSQL = MM_loginSQL & "," & MM_fldUserAuthorization
  MM_loginSQL = MM_loginSQL & " FROM [admin] WHERE Username = ? AND Passwd = ?"
  Set MM_rsUser_cmd = Server.CreateObject ("ADODB.Command")
  MM_rsUser_cmd.ActiveConnection = MM_connection_STRING
  MM_rsUser_cmd.CommandText = MM_loginSQL
  MM_rsUser_cmd.Parameters.Append MM_rsUser_cmd.CreateParameter("param1", 200, 1, 50, MM_valUsername) ' adVarChar
  MM_rsUser_cmd.Parameters.Append MM_rsUser_cmd.CreateParameter("param2", 200, 1, 50, Request.Form("Passwd")) ' adVarChar
  MM_rsUser_cmd.Prepared = true
  Set MM_rsUser = MM_rsUser_cmd.Execute

  If Not MM_rsUser.EOF Or Not MM_rsUser.BOF Then 
    ' username and password match - this is a valid user
    Session("MM_Username") = MM_valUsername
    If (MM_fldUserAuthorization <> "") Then
      Session("MM_UserAuthorization") = CStr(MM_rsUser.Fields.Item(MM_fldUserAuthorization).Value)
    Else
      Session("MM_UserAuthorization") = ""
    End If
    if CStr(Request.QueryString("accessdenied")) <> "" And false Then
      MM_redirectLoginSuccess = Request.QueryString("accessdenied")
    End If
    MM_rsUser.Close
    Response.Redirect(MM_redirectLoginSuccess)
  End If
  MM_rsUser.Close
  Response.Redirect(MM_redirectLoginFailed)
End If
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
    <td height="400" valign="top" bgcolor="#FFFFFF"><form ACTION="<%=MM_LoginAction%>" method="POST"  name="login">
      <p>&nbsp;</p>
      <table width="300" border="0" cellspacing="1" cellpadding="0" align="center" bgcolor="#000000">
        <tr bgcolor="#FFFFFF">
          <td><table width="100%" border="0" cellspacing="1" cellpadding="1">
              <tr>
                <td width="33%" align="right" height="30">用户名：</td>
                <td width="67%"><input name="UserName" maxlength="20" class="smallInput" size="20" />
                </td>
              </tr>
              <tr>
                <td width="33%" align="right" height="30">密 码：</td>
                <td width="67%"><input type="password" name="Passwd" maxlength="16" class="smallInput"
              size="20" />
                </td>
              </tr>
              <tr>
                <td colspan="2" height="15"></td>
              </tr>
          </table></td>
        </tr>
        <tr align="center" bgcolor="#FFFFFF">
          <td height="31"><input type="submit" name="Submit" value="确定" class="buttonface" />
            &nbsp;
            <input type="reset" name="Submit2" value="重写" class="buttonface" />
          </td>
        </tr>
      </table>
    </form></td>
  </tr>
</table>
<table width="770" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#6687BA">
  <tr>
    <td height="27" bgcolor="#FFFFFF"><div align="center" class="STYLE2">copyright&copy;new center </div></td>
  </tr>
</table>
</body>
</html>