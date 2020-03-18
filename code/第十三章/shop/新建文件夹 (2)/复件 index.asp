<%@LANGUAGE="VBSCRIPT"%>
<%
' *** Validate request to log in to this site.
MM_LoginAction = Request.ServerVariables("URL")
If Request.QueryString<>"" Then MM_LoginAction = MM_LoginAction + "?" + Server.HTMLEncode(Request.QueryString)
MM_valUsername=CStr(Request.Form("UserName"))
If MM_valUsername <> "" Then
  MM_fldUserAuthorization=""
  MM_redirectLoginSuccess="shop.asp"
  MM_redirectLoginFailed="link.html"
  MM_flag="ADODB.Recordset"
  set MM_rsUser = Server.CreateObject(MM_flag)
  MM_rsUser.ActiveConnection = MM_conn_STRING
  MM_rsUser.Source = "SELECT person_id, person_pass"
  If MM_fldUserAuthorization <> "" Then MM_rsUser.Source = MM_rsUser.Source & "," & MM_fldUserAuthorization
  MM_rsUser.Source = MM_rsUser.Source & " FROM person WHERE person_id='" & Replace(MM_valUsername,"'","''") &"' AND person_pass='" & Replace(Request.Form("Password"),"'","''") & "'"
  MM_rsUser.CursorType = 0
  MM_rsUser.CursorLocation = 2
  MM_rsUser.LockType = 3
  MM_rsUser.Open
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
<title>Freedom</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<script type="text/JavaScript">
<!--
function MM_swapImgRestore() { //v3.0
  var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
}

function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}

function MM_findObj(n, d) { //v4.01
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && d.getElementById) x=d.getElementById(n); return x;
}

function MM_swapImage() { //v3.0
  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
}
//-->
</script>
</head>

<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" text="#333333" onLoad="MM_preloadImages('image/bt_01_on.gif','image/bt_03_on.gif','image/bt_04_on.gif','image/bt_05_on.gif')" link="#6633FF" vlink="#6666FF" alink="#6666FF">
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
                <td width="346" valign="bottom"><table border="0" cellspacing="0" cellpadding="0">
                    <tr> 
                      <td><a href="../index.html" ><img src="image/bt_01_off.gif" width="86" height="18" border="0" name="Image5"></a></td>
                      <td><a href="javascript:;"><img src="image/bt_02_on.gif" width="86" height="18" border="0" name="Image1"></a></td>
                      <td><a href="shop.asp" ><img src="image/bt_03_off.gif" width="86" height="18" border="0" name="Image2"></a></td>
                      <td><a href="../html/chaxun1.asp" ><img src="image/bt_04_off.gif" width="86" height="18" border="0" name="Image3"></a></td>
                      <td><a href="link.html" ><img src="image/bt_05_off.gif" width="86" height="18" border="0" name="Image4"></a></td>
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
      <table width="640" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="16"><img src="image/transparent.gif" width="50" height="20"></td>
          <td width="618"><table border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td align="left" valign="middle"><table width="0" border="0" cellspacing="0" cellpadding="0">
                    <tr>
                      <td width="196">&nbsp;</td>
                    </tr>
                    <tr>
                      <td>&nbsp;</td>
                    </tr>
                  </table>
                    <img src="image/line.gif" width="590" height="11">
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
                          <td width="210" height=55 align="center"><input type="submit" name="Submit" value="提交">                          </td>
                          <td width="194" height=55 align="center"><input type="reset" name="Submit2" value="重置"></td>
                          <td width="184" align="center"><span class="STYLE1"><a href="zhuce.asp">注册</a></span></td>
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


