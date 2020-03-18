<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="Connections/connection.asp" -->
<%
' *** Validate request to log in to this site.
MM_LoginAction = Request.ServerVariables("URL")
If Request.QueryString <> "" Then MM_LoginAction = MM_LoginAction + "?" + Server.HTMLEncode(Request.QueryString)
MM_valUsername = CStr(Request.Form("name"))
If MM_valUsername <> "" Then
  Dim MM_fldUserAuthorization
  Dim MM_redirectLoginSuccess
  Dim MM_redirectLoginFailed
  Dim MM_loginSQL
  Dim MM_rsUser
  Dim MM_rsUser_cmd
  
  MM_fldUserAuthorization = ""
  MM_redirectLoginSuccess = "adminuser.asp"
  MM_redirectLoginFailed = "login.asp"

  MM_loginSQL = "SELECT usrname, passwd"
  If MM_fldUserAuthorization <> "" Then MM_loginSQL = MM_loginSQL & "," & MM_fldUserAuthorization
  MM_loginSQL = MM_loginSQL & " FROM passadmin WHERE usrname = ? AND passwd = ?"
  Set MM_rsUser_cmd = Server.CreateObject ("ADODB.Command")
  MM_rsUser_cmd.ActiveConnection = MM_connection_STRING
  MM_rsUser_cmd.CommandText = MM_loginSQL
  MM_rsUser_cmd.Parameters.Append MM_rsUser_cmd.CreateParameter("param1", 200, 1, 50, MM_valUsername) ' adVarChar
  MM_rsUser_cmd.Parameters.Append MM_rsUser_cmd.CreateParameter("param2", 200, 1, 50, Request.Form("pass")) ' adVarChar
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

<HTML>
<HEAD>
<TITLE>论坛管理登陆</TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=gb2312">
<style type="text/css">
<!--
body {
	margin: 0;
	overflow: hidden;
	scrollbar-face-color: D9E5F6;
	scrollbar-highlight-color: #FFFFFF;
	scrollbar-shadow-color: darkseablue;
	scrollbar-3dlight-color: D9E5F6;
	scrollbar-arrow-color: darkseablue;
	scrollbar-track-color: #f3faf4;
	scrollbar-darkshadow-color: #f3faf4;
}
td {
	font-size: 12px;
	line-height: 140%;
}
.copyright {
	padding-bottom: 10px;
}
.sysname {
	padding-bottom: 5px;
}
a:link {
	color: #000000;
	text-decoration: none;
}
a:visited {
	color: #000000;
	text-decoration: none;
}
a:hover {
	color: red;
	text-decoration: underline;
}
-->
</style>
</HEAD>
<body  text="#000000" leftmargin="0" topmargin="0" oncontextmenu=""return false;"">
<p>&nbsp;</p>
<table width="50%" border="0" cellspacing="1" cellpadding="3" align="center" bgcolor=#205E7B>
 <form name="form1" method="POST" action="<%=MM_LoginAction%>">
 <tr>
 <td colspan="2" height="25" align="center"><font color="#FFFFFF">管理登陆</font>
 </td>
 </tr>
 <tr>
 <td bgcolor=#64B3D9 align="center" width="40%">帐号： </td>
 <td bgcolor=#64B3D9 width="60%">
 <input type="text" name="name" size="16" value=>
 </td>
 </tr>
 <tr>
 <td bgcolor=#64B3D9 align="center" width="40%">密码： </td>
 <td bgcolor="#64B3D9" width="60%">
  <input  name=pass size="16" maxlength="15" type="password"></td></tr>
  <tr>
  <td colspan="2" align="center">
   <input  type=submit value="登 陆" name="submit"> 
    &nbsp;
    <input  type=reset value="取 消" name="submit2"></td>
    </tr></form>
</table>


</TABLE>
</BODY>
</HTML>
