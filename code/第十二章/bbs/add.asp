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
                  <a href="index.asp">��̳��ҳ</a>&nbsp;&nbsp;&nbsp;&nbsp;<a href="login.asp">�����½</a>
    &nbsp;</td>
  </tr>
</table>
<br />
<table width="98%" border="0" cellspacing="0" cellpadding="0" align="center"style="BORDER-LEFT: #000000 1px solid; BORDER-RIGHT: #000000 1px solid" bgcolor="#64B3D9" >
  <tr>
    <td><form ACTION="<%=MM_editAction%>" method="POST"  name="form1" id="form1">
          <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" bordercolorlight="#999999" bordercolordark="#FFFFFF" bgcolor="#205E7B" class="STYLE2">
            <tr>
              <td colspan="3" height="25" class="td1">����<span class="STYLE2">              ����������:</span></td>
            </tr>
            <tr>
              <td width="21%" align="center" bgcolor="#92C8E2">��&nbsp; ��&nbsp; �ˣ�</td>
              <td colspan="2" bgcolor="#FFFFFF"><input type="text" name="name" size="20" value="" />
                ��ID���� <b>5</b> �������ڣ�<font color="#FF0000">*</font> </td>
            </tr>
                    <tr> 
          <td width="16%" align="center" bgcolor="#92C8E2">�� �� �� �⣺</td>
          <td colspan="2" bgcolor="#92C8E2"> 
            <input type="text" name="title" size="50" maxlength="40" >
            ���������� <b>20</b> ���������ڣ�<font color="#FF0000">*</font></td>
        </tr>
            <tr>
              <td width="21%" align="center" bgcolor="#92C8E2">�� �� �� ����</td>
              <td width="19%" bgcolor="#92C8E2"><input type="text" name="email" size="20" />
              </td>
              <td width="60%" bgcolor="#92C8E2">�� ҳ �� ַ��
                <input type="text" name="url" size="30"  value="http://" />
              </td>
            </tr>
            <tr>
              <td width="21%" align="center" bgcolor="#DAEDF5"><img id="idface" src="images/01.gif" alt="�����������" /><br /><!-- wo d -->
                  <select name="face" size="1" onchange="document.images['idface'].src=options[selectedIndex].value;" >
                    <option selected="selected" value="images/01.gif">�û�ͷ��-01 </option>
                    <option selected="selected" value="images/02.gif">�û�ͷ��-02 </option>
                    <option selected="selected" value="images/03.gif">�û�ͷ��-03 </option>
                    <option selected="selected" value="images/04.gif">�û�ͷ��-04 </option>
                    <option selected="selected" value="images/05.gif">�û�ͷ��-05 </option>
                    <option selected="selected" value="images/06.gif">�û�ͷ��-06 </option>
                    <option selected="selected" value="images/07.gif">�û�ͷ��-07 </option>
                    <option selected="selected" value="images/08.gif">�û�ͷ��-08 </option>
                    <option selected="selected" value="images/09.gif">�û�ͷ��-09 </option>
                    <option selected="selected" value="images/10.gif">�û�ͷ��-10 </option>
                    <option selected="selected" value="images/11.gif">�û�ͷ��-11 </option>
                    <option selected="selected" value="images/12.gif">�û�ͷ��-12 </option>
                    <option selected="selected" value="images/13.gif">�û�ͷ��-13 </option>
                    <option selected="selected" value="images/14.gif">�û�ͷ��-14 </option>
                    <option selected="selected" value="images/15.gif">�û�ͷ��-15 </option>
                    <option selected="selected" value="images/16.gif">�û�ͷ��-16 </option>
                    <option selected="selected" value="images/17.gif">�û�ͷ��-17 </option>
                    <option selected="selected" value="images/18.gif">�û�ͷ��-18 </option>
                    <option selected="selected" value="images/19.gif">�û�ͷ��-19 </option>
                    <option selected="selected" value="images/20.gif">�û�ͷ��-20 </option>
                </select>
              </td>
              <td valign="top" colspan="2" bgcolor="#DAEDF5"><label>
                <textarea name="textarea" id="textarea" cols="100" rows="10"></textarea>
                </label>
              </td>
            </tr>
            <tr>
              <td colspan="3"><p align="center">
                  <input type="submit" value=" �ύ " name="B1" />
                &nbsp;
                <input type="reset" value=" ��� " name="B2" />
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