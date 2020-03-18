<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../Connections/connection.asp" -->
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
' *** Delete Record: construct a sql delete statement and execute it

If (CStr(Request("MM_delete")) = "form2" And CStr(Request("MM_recordId")) <> "") Then

  If (Not MM_abortEdit) Then
    ' execute the delete
    Set MM_editCmd = Server.CreateObject ("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_connection_STRING
    MM_editCmd.CommandText = "DELETE FROM yuding WHERE id = ?"
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 5, 1, -1, Request.Form("MM_recordId")) ' adDouble
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close

    ' append the query string to the redirect URL
    Dim MM_editRedirectUrl
    MM_editRedirectUrl = "dingdanmanage.asp"
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
Dim Recordset1
Dim Recordset1_cmd
Dim Recordset1_numRows

Set Recordset1_cmd = Server.CreateObject ("ADODB.Command")
Recordset1_cmd.ActiveConnection = MM_connection_STRING
Recordset1_cmd.CommandText = "SELECT * FROM yuding" 
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
<%
Dim MM_paramName 
%>
<%
' *** Go To Record and Move To Record: create strings for maintaining URL and Form parameters

Dim MM_keepNone
Dim MM_keepURL
Dim MM_keepForm
Dim MM_keepBoth

Dim MM_removeList
Dim MM_item
Dim MM_nextItem

' create the list of parameters which should not be maintained
MM_removeList = "&index="
If (MM_paramName <> "") Then
  MM_removeList = MM_removeList & "&" & MM_paramName & "="
End If

MM_keepURL=""
MM_keepForm=""
MM_keepBoth=""
MM_keepNone=""

' add the URL parameters to the MM_keepURL string
For Each MM_item In Request.QueryString
  MM_nextItem = "&" & MM_item & "="
  If (InStr(1,MM_removeList,MM_nextItem,1) = 0) Then
    MM_keepURL = MM_keepURL & MM_nextItem & Server.URLencode(Request.QueryString(MM_item))
  End If
Next

' add the Form variables to the MM_keepForm string
For Each MM_item In Request.Form
  MM_nextItem = "&" & MM_item & "="
  If (InStr(1,MM_removeList,MM_nextItem,1) = 0) Then
    MM_keepForm = MM_keepForm & MM_nextItem & Server.URLencode(Request.Form(MM_item))
  End If
Next

' create the Form + URL string and remove the intial '&' from each of the strings
MM_keepBoth = MM_keepURL & MM_keepForm
If (MM_keepBoth <> "") Then 
  MM_keepBoth = Right(MM_keepBoth, Len(MM_keepBoth) - 1)
End If
If (MM_keepURL <> "")  Then
  MM_keepURL  = Right(MM_keepURL, Len(MM_keepURL) - 1)
End If
If (MM_keepForm <> "") Then
  MM_keepForm = Right(MM_keepForm, Len(MM_keepForm) - 1)
End If

' a utility function used for adding additional parameters to these strings
Function MM_joinChar(firstItem)
  If (firstItem <> "") Then
    MM_joinChar = "&"
  Else
    MM_joinChar = ""
  End If
End Function
%>
<html>
<head>
<title>订单管理</title>
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
<table width="100%" height="606" border="0" cellpadding="0" cellspacing="0" background="../image/bg_top2.gif">
  <tr>
    <td height="89" align="left" valign="top" background="../image/bg_top2.gif">
     <table width="644" height="75" border="0" cellpadding="0" cellspacing="0">
       <tr>
         <td height="20" colspan="6">&nbsp;</td>
         <td width="104"><div align="right"><a href="login.asp"><img src="../image/con_contact.gif" border="0"></a></div></td>
       </tr>
        <tr>
           <td width="190" height="28"><img src="../image/obj_title.gif" width="190" height="28"></td>
           <td colspan="6">&nbsp;</td>
        </tr>
       <tr>
          <td>&nbsp;</td>
          <td width="86"><a href="../index.asp"><img src="../image/bt_01_off.gif" width="86" height="18" border="0"></a></td>
          <td width="86"><a href="../index.asp"><img src="../image/bt_03_off.gif" width="86" height="18" border="0"></a></td>
          <td width="83"><a href="../sousuo.asp"><img src="../image/bt_04_off.gif" width="86" height="18" border="0"></a></td>
          <td width="55"><a href="../dingdan.asp"><img src="../image/bt_08_off.gif" width="86" height="18" border="0"></a></td>
          <td width="40"><a href="../che.asp"><img src="../image/bt_09_off.gif" width="86" height="18" border="0"></a></td>
          <td><a href="mailto:fuping@sohu.co"><img src="../image/bt_05_off.gif" width="86" height="18" border="0"></a></td>
      </tr>
    </table>
   </td>
  </tr>
  <tr>
    <td height="154" align="left" valign="top">
    <table width="100%" height="520" border="0" cellpadding="0" cellspacing="0" bgcolor="#E7E7E7">
      <tr>
        <td height="71" valign="top"><% If Not Recordset1.EOF Or Not Recordset1.BOF Then %>
            <table width="80%" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#CCCC33">
              <tr>
                <td width="171"><div align="center">订单号</div></td>
                <td width="154"><div align="center">商品价格</div></td>
                <td width="171"><div align="center">订货时间</div></td>
                <td width="244"><div align="center">发货情况</div></td>
                <td width="98"><div align="center">删除</div></td>
              </tr>
              <% 
While ((Repeat1__numRows <> 0) AND (NOT Recordset1.EOF)) 
%>
                <tr>
                  <td><div align="center"><A HREF="../orderchakan.asp?<%= Server.HTMLEncode(MM_keepURL) & MM_joinChar(MM_keepURL) & "orderid=" & Recordset1.Fields.Item("orderid").Value %>"><%=(Recordset1.Fields.Item("orderid").Value)%></A></div></td>
                  <td><div align="center"><%=(Recordset1.Fields.Item("price").Value)%></div></td>
                  <td><div align="center"><%=(Recordset1.Fields.Item("yd_time").Value)%></div></td>
                  <td><form name="form1" method="post" action="saveresult.asp">
                      <label>
                      <input <%If (CStr((Recordset1.Fields.Item("yd_ifsuccess").Value)) = CStr("True")) Then Response.Write("checked=""checked""") : Response.Write("")%> type="checkbox" name="checkbox" id="checkbox">
                      已发货
                      <input name="id" type="hidden" value="<%=(Recordset1.Fields.Item("id").Value)%>" />
                      <input name="button" type="submit" id="button" value="提交" />
                      </label>
                  </form></td>
                  <td><form name="form2" method="POST" action="<%=MM_editAction%>">
                      <input type="submit" name="button2" id="button2" value="删除">
                      <input type="hidden" name="MM_delete" value="form2">
                      <input type="hidden" name="MM_recordId" value="<%= Recordset1.Fields.Item("id").Value %>">
                  </form></td>
                </tr>
                <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  Recordset1.MoveNext()
Wend
%>
            </table>
            <% End If ' end Not Recordset1.EOF Or NOT Recordset1.BOF %><br></td>
        </tr>
      <tr>
        <td height="449" align="left" valign="top"><div align="center">
          <% If Recordset1.EOF And Recordset1.BOF Then %>
            没有订单信息！
  <% End If ' end Recordset1.EOF And Recordset1.BOF %>
<br>
        </div></td>
        </tr>
    </table></td>
  </tr>
  <tr>
    <td height="38" background="../image/bg_bottom.gif"><img src="../image/transparent.gif" height="40"></td>
  </tr>
</table>
</body>
</html>
<%
Recordset1.Close()
Set Recordset1 = Nothing
%>
