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
If (CStr(Request("MM_insert")) = "Login") Then
  If (Not MM_abortEdit) Then
    ' execute the insert
    Dim MM_editCmd

    Set MM_editCmd = Server.CreateObject ("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_connection_STRING
    MM_editCmd.CommandText = "INSERT INTO [user] (user_id, user_name, user_pass, user_phone, user_address, user_email) VALUES (?, ?, ?, ?, ?, ?)" 
    MM_editCmd.Prepared = true
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 201, 1, -1, Request.Form("Userid")) ' adLongVarChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param2", 201, 1, -1, Request.Form("UserName")) ' adLongVarChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param3", 201, 1, -1, Request.Form("Password")) ' adLongVarChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param4", 201, 1, -1, Request.Form("phone")) ' adLongVarChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param5", 201, 1, -1, Request.Form("address")) ' adLongVarChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param6", 201, 1, -1, Request.Form("email")) ' adLongVarChar
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close

    ' append the query string to the redirect URL
    Dim MM_editRedirectUrl
    MM_editRedirectUrl = "index.asp"
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
<html>
<head>
<title>注册</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<script language="JavaScript">
<!--



function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}
//-->
</script>
<style type="text/css">
<!--
.STYLE1 {
	color: #0000FF;
	font-size: 14px;
}
-->
</style>
<script src="Scripts/AC_RunActiveContent.js" type="text/javascript"></script>
</head>

<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" text="#333333" link="#6633FF" vlink="#6666FF" alink="#6666FF">
<table width="100%" border="0" cellspacing="0" cellpadding="0" height="100%">
  <tr>
    <td height="361" valign="top"> 
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td background="image/bg_top2.gif" height="67" valign="top"> 
            <table width="536" border="0" cellspacing="0" cellpadding="0" background="">
              <tr> 
                <td width="190"><img src="image/transparent.gif" width="10" height="20"></td>
                <td width="346" align="right"><a href="../admin/login.asp"><img src="image/con_contact.gif" width="66" height="15" border="0"></a></td>
              </tr>
              <tr> 
                <td width="190"><img src="image/transparent.gif" width="10" height="22"></td>
                <td width="346">　</td>
              </tr>
              <tr>
                <td width="190">
                  <table border="0" cellspacing="0" cellpadding="0">
                    <tr>
                      <td><img src="image/obj_title.gif" width="190" height="28"></td>
                    </tr>
                    <tr>
                      <td><img src="image/transparent.gif" width="10" height="1"></td>
                    </tr>
                  </table>                </td>
                <td width="346" valign="bottom">
                  <table border="0" cellspacing="0" cellpadding="0">
                    <tr> 
                      <td><a href="index.asp"><img src="image/bt_01_off.gif" width="86" height="18" border="0" name="Image5"></a></td>
                      <td><a href="javascript:;"><img src="image/bt_02_on.gif" width="86" height="18" border="0" name="Image1"></a></td>
                      <td><a href="shop.asp"><img src="image/bt_03_off.gif" width="86" height="18" border="0" name="Image2"></a></td>
                      <td><a href="../html/chaxun1.asp"><img src="image/bt_04_off.gif" width="86" height="18" border="0" name="Image3"></a></td>
                      <td><a href="link.html"><img src="image/bt_05_off.gif" width="86" height="18" border="0" name="Image4"></a></td>
                      <td><img src="image/obj_bt.gif" width="9" height="18"></td>
                    </tr>
                  </table>                </td>
              </tr>
              <tr> 
                <td width="190"><img src="image/transparent.gif" width="10" height="25"></td>
                <td width="346">　</td>
              </tr>
            </table>          </td>
        </tr>
        <tr>
          <td height="5" valign="top"><img src="image/transparent.gif" width="20" height="10"></td>
        </tr>
      </table>
      <table width="618" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
          <td width="618"><table border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td align="left" valign="middle"><img src="image/line.gif" width="590" height="11">
                <table width="99%" height="136" align="center" cellpadding="0" cellspacing="0"bordercolor="#99BB99" style=" border-collapse: collapse">
                      <form ACTION="<%=MM_editAction%>" METHOD="POST" name="Login"> 
                        <tr>
                          <td width=210 height=33 align="right"><p  class="greenb">用户名：</td>
                      <td width="379" height=33 ><div align="left">
                              <input name="UserName"  type="text"  id="Userid" maxlength="20" >
                          </div></td>
                        </tr>
						                       <tr>
                          <td width=210 height=33 align="right"><p  class="greenb">真是姓名：</td>
                          <td height=33 ><div align="left">
                              <input name="Name"  type="text"  id="UserName" maxlength="20" >
                          </div></td>
                        </tr>
                        <tr>
                          <td width=210 height=46 align="right"><p  class="greenb">密　码：</td>
                          <td height=46 align="center"><div align="left">
                              <input name="Password"  type="password" maxlength="20" >
                          </div></td>
                        </tr>
						                        <tr>
                          <td width=210 height=33 align="right"><p  class="greenb">确认密码：</td>
                          <td height=33 ><div align="left">
                              <input name="querenpass"  type="password"  id="querenpass" maxlength="20" >
                          </div></td>
                        </tr>
                        <tr>
                          <td width=210 height=46 align="right"><p  class="greenb">电话：</td>
                          <td height=46 align="center"><div align="left">
                            <input name="phone" id="phone"  type="text" maxlength="20" >
                          </div></td>
                        </tr>
						                        <tr>
                          <td width=210 height=33 align="right"><p  class="greenb">地址：</td>
                          <td height=33 ><div align="left">
                              <input name="address"  type="text"  id="address" maxlength="20" >
                          </div></td>
                        </tr>
                        <tr>
                          <td width=210 height=46 align="right"><p  class="greenb">E-mail：</td>
                          <td height=46 align="center"><div align="left">
                              <input name="email" id="email"  type="text" maxlength="20" >
                          </div></td>
                        </tr>
						
                        <tr>
                          <td  height=55 align="center"><div align="right">
                             <input type="submit" name="Submit" value="提交">                             
                          </div></td>
                          <td height=55 align="center"><div align="left">
                            <input type="reset" name="Submit2" value="重置">                            
                          <span class="STYLE1"><a href="zhuce.asp"></a></span></div></td>
                        </tr>
                        <input type="hidden" name="MM_insert" value="Login">
                      </form>
                  </table>
                  <br>
                    <br></td>
              </tr>
          </table></td>
        </tr>
      </table>
    </td>
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
