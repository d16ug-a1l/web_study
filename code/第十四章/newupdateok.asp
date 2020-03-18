<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
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
If (CStr(Request("MM_update")) = "form1") Then
  If (Not MM_abortEdit) Then
    ' execute the update
    Dim MM_editCmd

    Set MM_editCmd = Server.CreateObject ("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_connection_STRING
    MM_editCmd.CommandText = "UPDATE [new] SET title = ?, typename = ?, tuijian = ?, picurl = ?, content = ?, Nfrom = ?, zuozhe = ?, selectpic = ? WHERE nid = ?" 
    MM_editCmd.Prepared = true
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 201, 1, -1, Request.Form("title")) ' adLongVarChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param2", 201, 1, -1, Request.Form("typeid")) ' adLongVarChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param3", 5, 1, -1, MM_IIF(Request.Form("checkbox1"), 1, 0)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param4", 201, 1, -1, Request.Form("picurl")) ' adLongVarChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param5", 201, 1, -1, Request.Form("txtcontent")) ' adLongVarChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param6", 201, 1, -1, Request.Form("Nfrom")) ' adLongVarChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param7", 201, 1, -1, Request.Form("zuozhe")) ' adLongVarChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param8", 5, 1, -1, MM_IIF(Request.Form("checkbox3"), 1, 0)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param9", 5, 1, -1, MM_IIF(Request.Form("MM_recordId"), Request.Form("MM_recordId"), null)) ' N/A
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close

    ' append the query string to the redirect URL
    Dim MM_editRedirectUrl
    MM_editRedirectUrl = "newupdate.asp"
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
Dim Recordset1__MMColParam
Recordset1__MMColParam = "118"
If (Request.QueryString("nid") <> "") Then 
  Recordset1__MMColParam = Request.QueryString("nid")
End If
%>
<%
Dim Recordset1
Dim Recordset1_cmd
Dim Recordset1_numRows

Set Recordset1_cmd = Server.CreateObject ("ADODB.Command")
Recordset1_cmd.ActiveConnection = MM_connection_STRING
Recordset1_cmd.CommandText = "SELECT * FROM new WHERE nid = ?" 
Recordset1_cmd.Prepared = true
Recordset1_cmd.Parameters.Append Recordset1_cmd.CreateParameter("param1", 5, 1, -1, Recordset1__MMColParam) ' adDouble

Set Recordset1 = Recordset1_cmd.Execute
Recordset1_numRows = 0
%>
<%
Dim Recordset2
Dim Recordset2_cmd
Dim Recordset2_numRows

Set Recordset2_cmd = Server.CreateObject ("ADODB.Command")
Recordset2_cmd.ActiveConnection = MM_connection_STRING
Recordset2_cmd.CommandText = "SELECT * FROM bankuai" 
Recordset2_cmd.Prepared = true

Set Recordset2 = Recordset2_cmd.Execute
Recordset2_numRows = 0
%><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>首页</title>
<link href="css.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
.STYLE2 {font-size: 10pt}
.STYLE9 {color: #000000}
-->
</style>
</head>
<body leftmargin="0" topmargin="0">
<!--#include file="head.asp" -->
<table width="770" border="0" align="center" cellpadding="1" bgcolor="#6687BA">
  <tr>
    <td height="400" valign="top" bgcolor="#FFFFFF"><p>&nbsp;</p>
      <form ACTION="<%=MM_editAction%>" METHOD="POST"  name="form1">
        <div align="center">
          <center>
            <table border="0" cellspacing="1" width="758" bordercolorlight="#000000" bordercolordark="#FFFFFF" cellpadding="0" bgcolor="#000000">
              <tr>
                <td width="100%" height="20" bgcolor="#FFFFFF"><p align="center"><b class="unnamed2 STYLE9">修 改 新 闻</b> </p></td>
              </tr>
              <tr align="center">
                <td width="100%"><table border="0" cellspacing="0" width="100%" cellpadding="0">
                    <tr>
                      <td width="17%" align="right" height="30" class="unnamed2" valign="middle" bgcolor="#FFFFFF"><div align="left">文章标题：</div></td>
                      <td width="83%" height="30" valign="middle" bgcolor="#FFFFFF">
                        <div align="left">
                          <input name="title" type="text" style="background-color:ffffff;color:000000;border: 1 double" value="<%=(Recordset1.Fields.Item("title").Value)%>" size="60" maxlength="100" />
                        </div></td>
                    </tr>
                    <tr>
                      <td width="17%" align="right" valign="middle" class="unnamed2" bgcolor="#FFFFFF"><div align="left">文章分类：</div></td>
             <td width="83%" valign="middle" bgcolor="#FFFFFF"><div align="left">
                          <select name="typeid" size="1" class="unnamed2">
                            <option value="<%=(Recordset2.Fields.Item("type").Value)%>" <%If (Not isNull((Recordset2.Fields.Item("type").Value))) Then If (CStr(Recordset2.Fields.Item("type").Value) = CStr((Recordset2.Fields.Item("type").Value))) Then Response.Write("selected=""selected""") : Response.Write("")%> ><%=(Recordset2.Fields.Item("type").Value)%></option>
                            <%
While (NOT Recordset2.EOF)
%>
                            <option value="<%=(Recordset2.Fields.Item("type").Value)%>" <%If (Not isNull((Recordset2.Fields.Item("type").Value))) Then If (CStr(Recordset2.Fields.Item("type").Value) = CStr((Recordset2.Fields.Item("type").Value))) Then Response.Write("selected=""selected""") : Response.Write("")%> ><%=(Recordset2.Fields.Item("type").Value)%></option>
<%
  Recordset2.MoveNext()
Wend
If (Recordset2.CursorType > 0) Then
  Recordset2.MoveFirst
Else
  Recordset2.Requery
End If
%>
                          </select>
                         </div></td>
                    </tr>
                    <tr>
                      <td width="17%" align="right" valign="middle" class="unnamed2" height="5" bgcolor="#FFFFFF"><div align="left">推荐新闻：</div></td>
                      <td width="83%" valign="middle" height="5" bgcolor="#FFFFFF"><div align="left">
                          <input <%If (Recordset1.Fields.Item("tuijian").Value= true) Then Response.Write("checked=""checked""") : Response.Write("")%> type="checkbox" name="checkbox1" value="1" class="unnamed5" />                      
                      </div></td>
                    </tr>

                    <tr>
                      <td width="17%" align="right" valign="middle" class="unnamed2" bgcolor="#FFFFFF"><div align="left">图片路径：</div></td>
                      <td width="83%" valign="middle" bgcolor="#FFFFFF"><div align="left">
                          <input name="picurl" type="text" style="background-color:ffffff;color:000000;border: 1 double" value="<%=(Recordset1.Fields.Item("picurl").Value)%>" />
                        </div></td>
                    </tr>
                    <tr>
                      <td width="17%" align="right" valign="middle" class="unnamed2" bgcolor="#FFFFFF"><div align="left">文章内容：</div></td>
                      <td width="83%" valign="middle" bgcolor="#FFFFFF"><div align="left">
                        <textarea rows="15" name="txtcontent" cols="70" class="smallarea"><%=(Recordset1.Fields.Item("content").Value)%></textarea>
                      </div></td>
                    </tr>

                    <tr>
                      <td width="17%" align="right" height="30" class="unnamed2" valign="middle" bgcolor="#FFFFFF"><div align="left">来源：&nbsp;</div></td>
                      <td width="83%" height="30" valign="middle" bgcolor="#FFFFFF"><div align="left">
                          <input name="Nfrom"  type="text" style="background-color:ffffff;color:000000;border: 1 double" value="<%=(Recordset1.Fields.Item("Nfrom").Value)%>" size="30" maxlength="100" />
                      </div></td>
                    </tr>
                    <tr>
                      <td width="17%" align="right" height="30" class="unnamed2" valign="middle" bgcolor="#FFFFFF"><div align="left">文章作者：</div></td>
                      <td width="83%" height="30" valign="middle" bgcolor="#FFFFFF"><div align="left">
                        <input name="zuozhe"  type="text" style="background-color:ffffff;color:000000;border: 1 double" value="<%=(Recordset1.Fields.Item("zuozhe").Value)%>" size="30" maxlength="100" />
                      </div></td>
                    </tr>

                    <tr>
                      <td width="17%" align="right" height="30" class="unnamed2" valign="middle" bgcolor="#FFFFFF"><div align="left">新闻是否含有图片</div></td>
                      <td width="83%" height="30" bgcolor="#FFFFFF" valign="middle"><div align="left">
                          <input <%If (Recordset1.Fields.Item("selectpic").Value= true) Then Response.Write("checked=""checked""") : Response.Write("")%> type="checkbox" name="checkbox3" value="1" class="unnamed5" />
                         
                      </div></td>
                    </tr>
                </table></td>
              </tr>
            </table>
          </center>
        </div>
        <div align="center">
          <center>
            <p>
              <input type="submit" value=" 修 改 "
  name="cmdok" class="unnamed5" />
              &nbsp;
              <input type="reset" value=" 清 除 "
  name="cmdcancel" class="unnamed5" />
            </p>
          </center>
        </div>
    
        
      
        <input type="hidden" name="MM_update" value="form1" />
        <input type="hidden" name="MM_recordId" value="<%= Recordset1.Fields.Item("nid").Value %>" />
      </form>      <p>&nbsp;</p></td>
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
<%
Recordset2.Close()
Set Recordset2 = Nothing
%>
