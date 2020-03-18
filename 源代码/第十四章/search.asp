<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Connections/connection.asp" -->
<%
Dim Recordset1__MMColParam
Recordset1__MMColParam = "1"
If (Request.Form("key") <> "") Then 
  Recordset1__MMColParam = Request.Form("key")
End If
%>
<%
Dim Recordset1
Dim Recordset1_cmd
Dim Recordset1_numRows

Set Recordset1_cmd = Server.CreateObject ("ADODB.Command")
Recordset1_cmd.ActiveConnection = MM_connection_STRING
Recordset1_cmd.CommandText = "SELECT * FROM new WHERE content LIKE ? ORDER BY dateandtime DESC" 
Recordset1_cmd.Prepared = true
Recordset1_cmd.Parameters.Append Recordset1_cmd.CreateParameter("param1", 200, 1, 255, "%" + Recordset1__MMColParam + "%") ' adVarChar

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
'  *** Recordset Stats, Move To Record, and Go To Record: declare stats variables

Dim Recordset1_total
Dim Recordset1_first
Dim Recordset1_last

' set the record count
Recordset1_total = Recordset1.RecordCount

' set the number of rows displayed on this page
If (Recordset1_numRows < 0) Then
  Recordset1_numRows = Recordset1_total
Elseif (Recordset1_numRows = 0) Then
  Recordset1_numRows = 1
End If

' set the first and last displayed record
Recordset1_first = 1
Recordset1_last  = Recordset1_first + Recordset1_numRows - 1

' if we have the correct record count, check the other stats
If (Recordset1_total <> -1) Then
  If (Recordset1_first > Recordset1_total) Then
    Recordset1_first = Recordset1_total
  End If
  If (Recordset1_last > Recordset1_total) Then
    Recordset1_last = Recordset1_total
  End If
  If (Recordset1_numRows > Recordset1_total) Then
    Recordset1_numRows = Recordset1_total
  End If
End If
%>

<%
' *** Recordset Stats: if we don't know the record count, manually count them

If (Recordset1_total = -1) Then

  ' count the total records by iterating through the recordset
  Recordset1_total=0
  While (Not Recordset1.EOF)
    Recordset1_total = Recordset1_total + 1
    Recordset1.MoveNext
  Wend

  ' reset the cursor to the beginning
  If (Recordset1.CursorType > 0) Then
    Recordset1.MoveFirst
  Else
    Recordset1.Requery
  End If

  ' set the number of rows displayed on this page
  If (Recordset1_numRows < 0 Or Recordset1_numRows > Recordset1_total) Then
    Recordset1_numRows = Recordset1_total
  End If

  ' set the first and last displayed record
  Recordset1_first = 1
  Recordset1_last = Recordset1_first + Recordset1_numRows - 1
  
  If (Recordset1_first > Recordset1_total) Then
    Recordset1_first = Recordset1_total
  End If
  If (Recordset1_last > Recordset1_total) Then
    Recordset1_last = Recordset1_total
  End If

End If
%>

<%
Dim Repeat2__numRows
Dim Repeat2__index

Repeat2__numRows = 10
Repeat2__index = 0
Recordset2_numRows = Recordset2_numRows + Repeat2__numRows
%>
<%
Dim Repeat3__numRows
Dim Repeat3__index

Repeat3__numRows = 10
Repeat3__index = 0
Recordset3_numRows = Recordset3_numRows + Repeat3__numRows
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>搜索结果</title>
<link href="css.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
.STYLE2 {font-size: 10pt}
.STYLE3 {color: #006699; font-size: 10pt; }
-->
</style>
</head>
<body leftmargin="0" topmargin="0">
<!--#include file="head.asp" -->
<table width="770" border="0" align="center" cellpadding="1" bgcolor="#6687BA">
  <tr>
    <td width="577" height="336" valign="top" bgcolor="#FFFFFF">共有<%=(Recordset1_total)%>条新闻<br />
      <table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#6687BA">
        <tr>
          <td bgcolor="#FFFFFF"><div align="right">
              <table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#6687BA">
                <tr>
                  <td height="20" background="images/bg11.gif"><span class="STYLE3">最近更新 </span></td>
                </tr>
                <tr>
                  <td bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                    <% If Not Recordset1.EOF Or Not Recordset1.BOF Then %>
                      <% 
While ((Repeat1__numRows <> 0) AND (NOT Recordset1.EOF)) 
%>
                        <tr>
                          <td><div align="left"><a href=xianshi.asp?nid=<%=(Recordset1.Fields.Item("nid").Value)%>><%=(Recordset1.Fields.Item("title").Value)%></a></div></td>
                          <td><div align="left"><%=(Recordset1.Fields.Item("dateandtime").Value)%></div></td>
                        </tr>
                        <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  Recordset1.MoveNext()
Wend
%>
                      <% End If ' end Not Recordset1.EOF Or NOT Recordset1.BOF %>

                                    </table>
                      <div align="right"><a href="liulan.asp"+<%= Request.QueryString("typeid") %>>&gt;&gt;更多新闻&gt;&gt;</a> </div></td>
                </tr>
              </table>
            <a href="liulan.asp"+<%= Request.QueryString("typeid") %>></a></div></td>
        </tr>
      </table>
      
      <div align="center">
        <% If Recordset1.EOF And Recordset1.BOF Then %>
          没有相关新闻，请重新查询！
  <% End If ' end Recordset1.EOF And Recordset1.BOF %>
</div></td>
    <td width="183" align="center" valign="top" bgcolor="#FFFFFF"><table width="164" border="0" cellpadding="0" cellspacing="1" bgcolor="#6687BA">
      <tr>
        <td height="20" background="images/bg11.gif"><div align="center"><strong><font color="#006699">站 内 搜 索</font></strong></div></td>
      </tr>
      <tr>
        <td height="88" valign="bottom" bgcolor="#FAFBFC"><form name="searchtitle" method="POST" action="search.asp" target="_blank">
            <div align="center">
              <input name="key" type="text"  value="请输入关键字" size="16">
              <br>
              <br>
              <input type="submit" name="Submit" value="搜 索">
              <input type="reset" name="Submit2" value="取 消" >
            </div>
        </form></td>
      </tr>
    </table>
    </td>
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