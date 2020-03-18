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
' IIf implementation
Function MM_IIf(condition, ifTrue, ifFalse)
  If condition = "" Then
    MM_IIf = ifFalse
  Else
    MM_IIf = ifTrue
  End If
End Function
%>
<%
If (CStr(Request("MM_insert")) = "form1") Then
  If (Not MM_abortEdit) Then
    ' execute the insert
    Dim MM_editCmd

    Set MM_editCmd = Server.CreateObject ("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_connection_STRING
    MM_editCmd.CommandText = "INSERT INTO yuding (orderid, yd_id, yd_name, yd_phone, yd_address, yd_time, price) VALUES (?, ?, ?, ?, ?, ?, ?)" 
    MM_editCmd.Prepared = true
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 5, 1, -1, MM_IIF(Request.Form("orderid"), Request.Form("orderid"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param2", 202, 1, 50, Request.Form("yd_id")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param3", 202, 1, 50, Request.Form("yd_name")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param4", 202, 1, 50, Request.Form("phone")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param5", 203, 1, 1073741823, Request.Form("address")) ' adLongVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param6", 135, 1, -1, MM_IIF(Request.Form("time"), Request.Form("time"), null)) ' adDBTimeStamp
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param7", 5, 1, -1, MM_IIF(Request.Form("price"), Request.Form("price"), null)) ' adDouble
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close

    ' append the query string to the redirect URL
    Dim MM_editRedirectUrl
    MM_editRedirectUrl = "dingdan.asp"
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
.STYLE2 {color: #0000FF;
	font-size: 14px;
}
-->
</style></head>

<body>
<table width="100%" height="413" border="0" cellpadding="0" cellspacing="0" background="image/bg_top2.gif">
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
    <td height="283" align="left" valign="top">
    <table width="100%" height="283" border="0" cellpadding="0" cellspacing="0" bgcolor="#E7E7E7">
      <tr>
        <td height="283" align="left" valign="top"><table width="640" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="16"><img src="image/transparent.gif" width="50" height="20"></td>
            <td width="618" valign="top"><table border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td><table width="583" border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td width="583">&nbsp;</td>
                      </tr>
                      <tr>
                        <td><form ACTION="<%=MM_editAction%>" METHOD="POST" name="form1">
                            <table width="630" align="center" cellpadding="0" cellspacing="0"bordercolor="#99BB99" style=" border-collapse: collapse">
                              <tr>
                                <td width=186 height="29" align="right" class="STYLE2"><p align="center"  class="greenb">订单编号                                                  
                                </td>
                                <td width=442 align="right"><div align="left">
                                    <input type="text" name="orderid" value="<%= Session("orderid") %>">
                                </div></td>
                              </tr>
                              <tr>
                                <td height="27" align="right" class="STYLE2"><div align="center">用户名</div></td>
                                <td align="right"><div align="left">
                                    <input name="yd_id" type="text" value="<%= Session("MM_Username") %>">
                                </div></td>
                              </tr>
                              <tr>
                                <td height="27" align="right" class="STYLE2"><div align="center">真实姓名</div></td>
                                <td align="right"><div align="left">
                                    <input type="text" name="yd_name">
                                </div></td>
                              </tr>
                              <tr>
                                <td height="33" align="right" class="STYLE2"><div align="center"><span class="greenb">电话</span> </div></td>
                                <td align="right"><div align="left">
                                  <input type="text" name="phone">
                                </div></td>
                              </tr>
                              <tr>
                                <td height="23" align="right" class="STYLE2"><div align="center"><span class="greenb">地址</span></div></td>
                                <td align="right"><div align="left">
                                  <input type="text" name="address">
                                </div></td>
                              </tr>
                              <tr>
                                <td height="24" align="right" class="STYLE2"><div align="center"><span class="greenb">付款方式</span><span class="greenb"></span></div></td>
                                <td align="right"><div align="left">
                                    <select name="fangshi" size="1">
                                      <option value="汇款">汇款</option>
                                      <option value="信用卡">信用卡</option>
                                    </select>
                                    </label>
                                </div></td>
                              </tr>
                              <tr>
                                <td width="186" height="28" align="center"><div align="center"><span class="STYLE2"><span class="greenb">上传时间</span></span></div></td>
                                <td width="442" align="center"><div align="left">
                                    <input name="time" type="text" value="<%=now()%>">
                                </div></td>
                              </tr>
                                                            <tr>
                                <td width="186" height="28" align="center"><div align="center">总价格</div></td>
                                <td width="442" align="center"><div align="left">
                                    <input name="price" type="text" value="<%= Session("price") %>">
                                </div></td>
                              </tr>
                              <tr>
                                <td width="186" height="32" align="center"><div align="center"><span class="STYLE2"><a href="zhuce.asp">
                                    <label>
                                    <input type="submit" name="Submit" value="提交">
                                    </label>
                                </a></span></div></td>
                                <td width="442" align="center"><div align="left">
                                    <label>
                                    <input type="reset" name="Submit2" value="重置">
                                    </label>
                                </div></td>
                              </tr>
                            </table>
                          
<input type="hidden" name="MM_insert" value="form1">
                        </form></td>
                      </tr>
                  </table></td>
                </tr>
            </table></td>
          </tr>
        </table>
          <br></td>
        </tr>
    </table></td>
  </tr>
  <tr>
    <td height="38" background="image/bg_bottom.gif"><img src="image/transparent.gif" height="40"></td>
  </tr>
</table>
</body>
</html>