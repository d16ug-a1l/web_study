<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Connections/connection.asp" -->
<%
Dim Recordset1__MMColParam
Recordset1__MMColParam = "1"
If (Request.Form("keyword") <> "") Then 
  Recordset1__MMColParam = Request.Form("keyword")
End If
%>
<%
Dim Recordset1
Dim Recordset1_cmd
Dim Recordset1_numRows

Set Recordset1_cmd = Server.CreateObject ("ADODB.Command")
Recordset1_cmd.ActiveConnection = MM_connection_STRING
Recordset1_cmd.CommandText = "SELECT * FROM postMain WHERE main_subject LIKE ? ORDER BY main_time DESC" 
Recordset1_cmd.Prepared = true
Recordset1_cmd.Parameters.Append Recordset1_cmd.CreateParameter("param1", 200, 1, 50, "%" + Recordset1__MMColParam + "%") ' adVarChar

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
<title>搜索结果</title>
</head>

<body>
<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0" >
  <tr> <td></td>
                <td  height=37 align=right valign="bottom" background="images/top.jpg" bgcolor="#6EB7DA" class="bg"><br>
                  <br>
                  <a href="index.asp">论坛首页</a>&nbsp;&nbsp;&nbsp;&nbsp;<a href="login.asp">管理登陆</a>
    &nbsp;</td>
  </tr>
</table>
<table border="0" width="98%" cellspacing="1" cellpadding="0"  bgcolor="#205E7B" align="center" >
        <tr> 
          <td colspan="2" align="center" height="25"><b><font color="#FFFFFF">帖子主题</font></b></td>
          <td align="center" width="179"><b><font color="#FFFFFF">作&nbsp;者</font></b></td>
          <td align="center" width="60"><b><font color="#FFFFFF">类型</font></b></td>
          <td align="center" width="323" ><b><font color="#FFFFFF">发表时间</font></b></td>
  </tr>
        <tr>
        </tr>

</table>
<table border="1" width="98%" cellspacing="1" cellpadding="0"  bgcolor="#CCFFFF" align="center" ><tr><td colspan="2" align="center" height="25"><table width="100%" border="1" align="center" cellpadding="0" cellspacing="1"  bgcolor="#CCFFFF" >
  <% If Not Recordset1.EOF Or Not Recordset1.BOF Then %>
    <% 
While ((Repeat1__numRows <> 0) AND (NOT Recordset1.EOF)) 
%>
        <tr>
          <td colspan="2" align="center" height="25"><b><font color="#000000"><a href=show.asp?main_id=<%=(Recordset1.Fields.Item("main_id").Value)%>><%=(Recordset1.Fields.Item("main_subject").Value)%></a></font></b></td>
          <td align="center" width="176"><b><font color="#000000"><%=(Recordset1.Fields.Item("main_name").Value)%></font></b></td>
          <td align="center" width="59"><img src="<%=(Recordset1.Fields.Item("main_important").Value)%>" vspace="2" /></td>
          <td align="center" width="317" ><%=(Recordset1.Fields.Item("main_time").Value)%></td>
        </tr>
        <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  Recordset1.MoveNext()
Wend
%>
    <% End If ' end Not Recordset1.EOF Or NOT Recordset1.BOF %>
</table>  <b><font color="#000000"></font></b></td>
          </tr>
          <tr>
          <td><div align="center">
            <% If Recordset1.EOF And Recordset1.BOF Then %>
              没有相关帖子，请重新查询！
  <% End If ' end Recordset1.EOF And Recordset1.BOF %>
</div></td>
          </tr>
</table>
</body>
</html>
<%
Recordset1.Close()
Set Recordset1 = Nothing
%>