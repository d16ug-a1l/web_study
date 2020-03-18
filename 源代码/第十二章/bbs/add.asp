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
If (CStr(Request("MM_insert")) = "form1") Then
  If (Not MM_abortEdit) Then
    ' execute the insert
    Dim MM_editCmd

    Set MM_editCmd = Server.CreateObject ("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_connection_STRING
    MM_editCmd.CommandText = "INSERT INTO postMain (main_name, main_subject, main_email, main_url, main_face, main_content) VALUES (?, ?, ?, ?, ?, ?)" 
    MM_editCmd.Prepared = true
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 202, 1, 50, Request.Form("name")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param2", 202, 1, 50, Request.Form("title")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param3", 202, 1, 50, Request.Form("email")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param4", 202, 1, 50, Request.Form("url")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param5", 202, 1, 50, Request.Form("face")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param6", 203, 1, 1073741823, Request.Form("textarea")) ' adLongVarWChar
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
<style type="text/css">
<!--
.bg {
	background-repeat: no-repeat;
}
.STYLE1 {font-size: 12px}
.STYLE2 {font-size: 14px}
-->
</style>
<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0" >
  <tr> <td></td>
                <td  height=37 align=right valign="bottom" background="images/top.jpg" bgcolor="#6EB7DA" class="bg"><br>
                  <br>
                  <a href="index.asp">论坛首页</a>&nbsp;&nbsp;&nbsp;&nbsp;<a href="login.asp">管理登陆</a>
    &nbsp;</td>
  </tr>
</table>
<br />
<table width="98%" border="0" cellspacing="0" cellpadding="0" align="center"style="BORDER-LEFT: #000000 1px solid; BORDER-RIGHT: #000000 1px solid" bgcolor="#64B3D9" >
  <tr>
    <td><form ACTION="<%=MM_editAction%>" method="POST"  name="form1" id="form1">
          <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" bordercolorlight="#999999" bordercolordark="#FFFFFF" bgcolor="#205E7B" class="STYLE2">
            <tr>
              <td colspan="3" height="25" class="td1">　　<span class="STYLE2">              发表新主题:</span></td>
            </tr>
            <tr>
              <td width="21%" align="center" bgcolor="#92C8E2">发&nbsp; 表&nbsp; 人：</td>
              <td colspan="2" bgcolor="#FFFFFF"><input type="text" name="name" size="20" value="" />
                （ID限制 <b>5</b> 汉字以内）<font color="#FF0000">*</font> </td>
            </tr>
                    <tr> 
          <td width="16%" align="center" bgcolor="#92C8E2">帖 子 主 题：</td>
          <td colspan="2" bgcolor="#92C8E2"> 
            <input type="text" name="title" size="50" maxlength="40" >
            （标题限制 <b>20</b> 个汉字以内）<font color="#FF0000">*</font></td>
        </tr>
            <tr>
              <td width="21%" align="center" bgcolor="#92C8E2">电 子 邮 件：</td>
              <td width="19%" bgcolor="#92C8E2"><input type="text" name="email" size="20" />
              </td>
              <td width="60%" bgcolor="#92C8E2">主 页 地 址：
                <input type="text" name="url" size="30"  value="http://" />
              </td>
            </tr>
            <tr>
              <td width="21%" align="center" bgcolor="#DAEDF5"><img id="idface" src="images/01.gif" alt="个人形象代表" /><br /><!-- wo d -->
                  <select name="face" size="1" onchange="document.images['idface'].src=options[selectedIndex].value;" >
                    <option selected="selected" value="images/01.gif">用户头像-01 </option>
                    <option selected="selected" value="images/02.gif">用户头像-02 </option>
                    <option selected="selected" value="images/03.gif">用户头像-03 </option>
                    <option selected="selected" value="images/04.gif">用户头像-04 </option>
                    <option selected="selected" value="images/05.gif">用户头像-05 </option>
                    <option selected="selected" value="images/06.gif">用户头像-06 </option>
                    <option selected="selected" value="images/07.gif">用户头像-07 </option>
                    <option selected="selected" value="images/08.gif">用户头像-08 </option>
                    <option selected="selected" value="images/09.gif">用户头像-09 </option>
                    <option selected="selected" value="images/10.gif">用户头像-10 </option>
                    <option selected="selected" value="images/11.gif">用户头像-11 </option>
                    <option selected="selected" value="images/12.gif">用户头像-12 </option>
                    <option selected="selected" value="images/13.gif">用户头像-13 </option>
                    <option selected="selected" value="images/14.gif">用户头像-14 </option>
                    <option selected="selected" value="images/15.gif">用户头像-15 </option>
                    <option selected="selected" value="images/16.gif">用户头像-16 </option>
                    <option selected="selected" value="images/17.gif">用户头像-17 </option>
                    <option selected="selected" value="images/18.gif">用户头像-18 </option>
                    <option selected="selected" value="images/19.gif">用户头像-19 </option>
                    <option selected="selected" value="images/20.gif">用户头像-20 </option>
                </select>
              </td>
              <td valign="top" colspan="2" bgcolor="#DAEDF5"><label>
                <textarea name="textarea" id="textarea" cols="100" rows="10"></textarea>
                </label>
              </td>
            </tr>
            <tr>
              <td colspan="3"><p align="center">
                  <input type="submit" value=" 提交 " name="B1" />
                &nbsp;
                <input type="reset" value=" 清除 " name="B2" />
              </p></td>
            </tr>
          </table>
        <br />
        <input type="hidden" name="MM_insert" value="form1" />
    </form>
    </td>
  </tr>
</table>
</body>
</html>