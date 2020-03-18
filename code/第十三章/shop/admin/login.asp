<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../Connections/connection.asp" -->
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
  MM_redirectLoginSuccess = "index.asp"
  MM_redirectLoginFailed = "login.asp"

  MM_loginSQL = "SELECT username, password"
  If MM_fldUserAuthorization <> "" Then MM_loginSQL = MM_loginSQL & "," & MM_fldUserAuthorization
  MM_loginSQL = MM_loginSQL & " FROM [admin] WHERE username = ? AND password = ?"
  Set MM_rsUser_cmd = Server.CreateObject ("ADODB.Command")
  MM_rsUser_cmd.ActiveConnection = MM_connection_STRING
  MM_rsUser_cmd.CommandText = MM_loginSQL
  MM_rsUser_cmd.Parameters.Append MM_rsUser_cmd.CreateParameter("param1", 200, 1, 50, MM_valUsername) ' adVarChar
  MM_rsUser_cmd.Parameters.Append MM_rsUser_cmd.CreateParameter("param2", 200, 1, 50, Request.Form("Password")) ' adVarChar
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
    if CStr(Request.QueryString("accessdenied")) <> "" And true Then
      MM_redirectLoginSuccess = Request.QueryString("accessdenied")
    End If
    MM_rsUser.Close
    Response.Redirect(MM_redirectLoginSuccess)
  End If
  MM_rsUser.Close
  Response.Redirect(MM_redirectLoginFailed)
End If
%>
<html>
<head>
<title>Freedom</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type="text/css">
<!--
.STYLE1 {
	color: #0000FF;
	font-size: 14px;
}
-->
</style>
</head>

<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" text="#333333" >
<table width="100%" border="0" cellspacing="0" cellpadding="0" height="100%">
  <tr>
    <td height="361" valign="top"><table width="640" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
          <td width="16">&nbsp;</td>
          <td width="618"><table border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td align="left" valign="middle"><table width="579" border="0" cellspacing="0" cellpadding="0">
                    <tr>
                      <td width="293">&nbsp;</td>
                    </tr>
                    <tr>
                      <td>&nbsp;</td>
                    </tr>
                  </table>
                    <table width="100%" height="197" align="center" cellpadding="0" cellspacing="0"bordercolor="#99BB99" style=" border-collapse: collapse">
                      <form ACTION="<%=MM_LoginAction%>" METHOD="POST" name="Login"> 
                        <tr>
                          <td width=210 height=59 align="right"><p  class="greenb">用户名：</td>
                          <td height=59 colspan="2" ><div align="left">
                              <input name="UserName"  type="text"  id="UserName4" maxlength="20" >
                          </div></td>
                        </tr>
                        <tr>
                          <td width=210 height=81 align="right"><p  class="greenb">密　码：</td>
                          <td height=81 colspan="2" align="center"><div align="left">
                              <input name="Password"  type="password" maxlength="20" >
                          </div></td>
                        </tr>
                        <tr>
                          <td width="210" height=55 align="center"><div align="right">
                            <input type="submit" name="Submit" value="提交">                          
                          </div></td>
                          <td width="194" height=55 align="center"><input type="reset" name="Submit2" value="重置"></td>
                          <td width="184" align="center"><span class="STYLE1"><a href="zhuce.asp"></a></span></td>
                        </tr>
                      </form>
                    </table>
                  <br>
                    <br></td>
              </tr>
          </table></td>
        </tr>
      </table></td>
  </tr>
  <tr>
    <td valign="bottom"> 
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td background="image/bg_bottom.gif" align="right">
            <table border="0" cellspacing="0" cellpadding="0" background="">
              <tr> 
                <td height="14"><img src="image/transparent.gif" width="210" height="20"></td>
                <td rowspan="2"><img src="image/transparent.gif" width="3" height="40"></td>
              </tr>
            </table>          </td>
        </tr>
      </table>    </td>
  </tr>
</table>
</body>
</html>
