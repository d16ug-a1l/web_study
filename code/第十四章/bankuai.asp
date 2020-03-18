<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Connections/connection.asp" -->
<% If  (Session("MM_Username")="")Then
response.Redirect("login.asp")
 End If 
 %>
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
    MM_editCmd.CommandText = "INSERT INTO bankuai (type) VALUES (?)" 
    MM_editCmd.Prepared = true
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 202, 1, 20, Request.Form("bk")) ' adVarWChar
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close

    ' append the query string to the redirect URL
    Dim MM_editRedirectUrl
    MM_editRedirectUrl = "bankuai.asp"
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
Recordset1_cmd.CommandText = "SELECT * FROM bankuai" 
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
    <td height="400" valign="top" bgcolor="#FFFFFF"><form id="form1" name="form1" method="POST" action="<%=MM_editAction%>">
      <label>
      <div align="center">
        <input type="text" name="bk" id="bk" />
        &nbsp;&nbsp;
        <input type="submit" name="button" id="button" value="提交" />
      </div>
      
            <input type="hidden" name="MM_insert" value="form1" />
</label>
                        </form>
      <table width="80%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#0033FF">
        <tr>
          <td width="40%" bgcolor="#FFFFFF"><div align="center">版块名称</div></td>
          <td width="35%" bgcolor="#FFFFFF"><div align="center">修改</div></td>
          <td width="25%" bgcolor="#FFFFFF"><div align="center">删除</div></td>
        </tr>
        <% 
While ((Repeat1__numRows <> 0) AND (NOT Recordset1.EOF)) 
%>
          <tr>
  
            <form  method="post" action="bkxiugai.asp">
             <td height="28" bgcolor="#FFFFFF">         
                <input name="type" type="text" id="textfield" value="<%=(Recordset1.Fields.Item("type").Value)%>" />            </td>
              <td bgcolor="#FFFFFF"><div align="center">
                <div align="center">
                  <input type="submit" name="button2" id="button2" value="修改" />
                  <input type="hidden"  name="typeid"value="<%=(Recordset1.Fields.Item("typeid").Value)%>" />
                </div>
                <div align="center">
                  <label></label>
                </div></td>   </form>
            <td bgcolor="#FFFFFF"><div align="center"><a href="bkshanchu.asp?typeid=<%=(Recordset1.Fields.Item("typeid").Value)%>">删除</a></div></td>
          </tr>
          <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  Recordset1.MoveNext()
Wend
%>
      </table>
      <p align="center" class="STYLE6"><br />
    </p>
    </p></td>
  </tr>
</table>
<table width="770" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#6687BA">
  <tr>
    <td height="27" bgcolor="#FFFFFF"><div align="center" class="STYLE2">copyright&copy;new center </div></td>
  </tr>
</table>
</body>
</html>
<%
Recordset1.Close()
Set Recordset1 = Nothing
%>
