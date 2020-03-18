<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="Connections/connection.asp" -->
<%
Dim Recordset1
Dim Recordset1_cmd
Dim Recordset1_numRows

Set Recordset1_cmd = Server.CreateObject ("ADODB.Command")
Recordset1_cmd.ActiveConnection = MM_connection_STRING
Recordset1_cmd.CommandText = "SELECT * FROM shop ORDER BY id DESC" 
Recordset1_cmd.Prepared = true

Set Recordset1 = Recordset1_cmd.Execute
Recordset1_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = 10
Repeat1__index = 0
Recordset1_numRows = Recordset1_numRows + Repeat1__numRows
%>
<%
' *** Validate request to log in to this site.
MM_LoginAction = Request.ServerVariables("URL")
If Request.QueryString <> "" Then MM_LoginAction = MM_LoginAction + "?" + Server.HTMLEncode(Request.QueryString)
MM_valUsername = CStr(Request.Form("username"))
If MM_valUsername <> "" Then
  Dim MM_fldUserAuthorization
  Dim MM_redirectLoginSuccess
  Dim MM_redirectLoginFailed
  Dim MM_loginSQL
  Dim MM_rsUser
  Dim MM_rsUser_cmd
  
  MM_fldUserAuthorization = ""
  MM_redirectLoginSuccess = "shop.asp"
  MM_redirectLoginFailed = "error.asp"

  MM_loginSQL = "SELECT user_id, user_pass"
  If MM_fldUserAuthorization <> "" Then MM_loginSQL = MM_loginSQL & "," & MM_fldUserAuthorization
  MM_loginSQL = MM_loginSQL & " FROM [user] WHERE user_id = ? AND user_pass = ?"
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
<html>
<head>
<title>网上购物网站</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">

<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
}
a {
	font-family: 宋体;
}
a:link {
	text-decoration: none;
}
a:visited {
	text-decoration: none;
}
a:hover {
	text-decoration: none;
}
a:active {
	text-decoration: none;
}
.STYLE1 {
	color: #0000FF;
	font-size: 36px;
}
-->
</style></head>

<body>
<table width="100%" height="606" border="0" cellpadding="0" cellspacing="0" background="image/bg_top2.gif">
  <tr>
    <td height="89" align="left" valign="top" background="image/bg_top2.gif">
     <table width="644" height="75" border="0" cellpadding="0" cellspacing="0">
       <tr>
         <td height="20" colspan="6">&nbsp;</td>
         <td width="104"><div align="right"><a href="admin/login.asp"><img src="image/con_contact.gif" border="0"></a></div></td>
       </tr>
        <tr>
           <td width="190" height="28"><img src="image/obj_title.gif" width="190" height="28"></td>
           <td colspan="6">&nbsp;</td>
        </tr>
       <tr>
          <td>&nbsp;</td>
          <td width="86"><a href="index.asp"><img src="image/bt_01_off.gif" width="86" height="18" border="0"></a></td>
          <td width="86"><a href="index.asp"><img src="image/bt_03_off.gif" width="86" height="18" border="0"></a></td>
          <td width="83"><a href="sousuo.asp"><img src="image/bt_04_off.gif" width="86" height="18" border="0"></a></td>
          <td width="55"><a href="dingdan.asp"><img src="image/bt_08_off.gif" width="86" height="18" border="0"></a></td>
          <td width="40"><a href="che.asp"><img src="image/bt_09_off.gif" width="86" height="18" border="0"></a></td>
          <td><a href="mailto:fuping@sohu.co"><img src="image/bt_05_off.gif" width="86" height="18" border="0"></a></td>
      </tr>
    </table>
   </td>
  </tr>
  <tr>
    <td height="154" align="left" valign="top">
    <table width="100%" height="476" border="0" cellpadding="0" cellspacing="0" bgcolor="#E7E7E7">
      <tr>
        <td height="27">
          <form ACTION="<%=MM_LoginAction%>" name="form1" method="POST">
          用户名：
          <label>
          <input type="text" name="username" id="username">
          </label>
                密码：
                <label>
                <input type="password" name="pass" id="pass">
                </label>
                <label>
                <input type="submit" name="button" id="button" value="提交">
                </label>
                <label>
                <input type="reset" name="button2" id="button2" value="重置">
                </label>
                <a href="zhuce.asp">新用户注册                </a>
            </form>          </td>
        </tr>
      <tr>
        <td height="449" align="left" valign="top">
        <span class="STYLE1">最新商品浏览</span> <br><br>
        <% 
While ((Repeat1__numRows <> 0) AND (NOT Recordset1.EOF)) 
%>
          <table width="826" border="0" cellspacing="0" cellpadding="0">
              <tr>
                  <td width="214" rowspan="5"><img src="<%=(Recordset1.Fields.Item("sp_pic").Value)%>" width="150" height="100" alt=""></td>
                 
                  <td width="91" height="29">商品名称：</td>
                  <td width="194"><%=(Recordset1.Fields.Item("sp_title").Value)%></td>
                  <td width="111" valign="bottom">商品价格：</td>
                  <td width="216" valign="bottom"><%=(Recordset1.Fields.Item("sp_price").Value)%></td>
              </tr>
              <tr>
                  <td height="31">商品类别：</td>
                  <td><%=(Recordset1.Fields.Item("sp_leibie").Value)%></td>
                  <td valign="bottom">剩余数量：</td>
                  <td valign="bottom"><%=(Recordset1.Fields.Item("sp_num").Value)%></td>
              </tr>
            <tr>
              <td height="32">商品说明：</td>
                  <td colspan="2"><%=(Recordset1.Fields.Item("sp_shuoming").Value)%></td>
                  <td><div align="right"><a href=addche.asp?id=<%=(Recordset1.Fields.Item("id").Value)%>>购买</a></div></td>
            </tr>
                  </table>
          <hr>
          <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  Recordset1.MoveNext()
Wend
%></td>
        </tr>
    </table></td>
  </tr>
  <tr>
    <td height="38" background="image/bg_bottom.gif"><img src="image/transparent.gif" height="40"></td>
  </tr>
</table>
</body>
</html>
<%
Recordset1.Close()
Set Recordset1 = Nothing
%>
