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
    MM_editCmd.CommandText = "INSERT INTO postRe (re_subject, m_id, re_name, re_email, re_url, re_face, re_content) VALUES (?, ?, ?, ?, ?, ?, ?)" 
    MM_editCmd.Prepared = true
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 202, 1, 50, Request.Form("subject")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param2", 5, 1, -1, MM_IIF(Request.Form("id"), Request.Form("id"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param3", 202, 1, 50, Request.Form("name")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param4", 202, 1, 50, Request.Form("email")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param5", 202, 1, 50, Request.Form("url")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param6", 202, 1, 50, Request.Form("face")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param7", 203, 1, 1073741823, Request.Form("textarea")) ' adLongVarWChar
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close

    ' append the query string to the redirect URL
    Dim MM_editRedirectUrl
    MM_editRedirectUrl = "show.asp"
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
Recordset1__MMColParam = "1"
If (Request.QueryString("main_id") <> "") Then 
  Recordset1__MMColParam = Request.QueryString("main_id")
End If
%>
<%
Dim Recordset1
Dim Recordset1_cmd
Dim Recordset1_numRows

Set Recordset1_cmd = Server.CreateObject ("ADODB.Command")
Recordset1_cmd.ActiveConnection = MM_connection_STRING
Recordset1_cmd.CommandText = "SELECT * FROM postMain WHERE main_id = ?" 
Recordset1_cmd.Prepared = true
Recordset1_cmd.Parameters.Append Recordset1_cmd.CreateParameter("param1", 5, 1, -1, Recordset1__MMColParam) ' adDouble

Set Recordset1 = Recordset1_cmd.Execute
Recordset1_numRows = 0
%>
<%
Dim Recordset2__MMColParam
Recordset2__MMColParam = "1"
If (Request.QueryString("main_id") <> "") Then 
  Recordset2__MMColParam = Request.QueryString("main_id")
End If
%>
<%
Dim Recordset2
Dim Recordset2_cmd
Dim Recordset2_numRows

Set Recordset2_cmd = Server.CreateObject ("ADODB.Command")
Recordset2_cmd.ActiveConnection = MM_connection_STRING
Recordset2_cmd.CommandText = "SELECT * FROM postRe WHERE m_id = ? ORDER BY re_time DESC" 
Recordset2_cmd.Prepared = true
Recordset2_cmd.Parameters.Append Recordset2_cmd.CreateParameter("param1", 5, 1, -1, Recordset2__MMColParam) ' adDouble

Set Recordset2 = Recordset2_cmd.Execute
Recordset2_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = 5
Repeat1__index = 0
Recordset2_numRows = Recordset2_numRows + Repeat1__numRows
%>
<%
'  *** Recordset Stats, Move To Record, and Go To Record: declare stats variables

Dim Recordset2_total
Dim Recordset2_first
Dim Recordset2_last

' set the record count
Recordset2_total = Recordset2.RecordCount

' set the number of rows displayed on this page
If (Recordset2_numRows < 0) Then
  Recordset2_numRows = Recordset2_total
Elseif (Recordset2_numRows = 0) Then
  Recordset2_numRows = 1
End If

' set the first and last displayed record
Recordset2_first = 1
Recordset2_last  = Recordset2_first + Recordset2_numRows - 1

' if we have the correct record count, check the other stats
If (Recordset2_total <> -1) Then
  If (Recordset2_first > Recordset2_total) Then
    Recordset2_first = Recordset2_total
  End If
  If (Recordset2_last > Recordset2_total) Then
    Recordset2_last = Recordset2_total
  End If
  If (Recordset2_numRows > Recordset2_total) Then
    Recordset2_numRows = Recordset2_total
  End If
End If
%>
<%
' *** Recordset Stats: if we don't know the record count, manually count them

If (Recordset2_total = -1) Then

  ' count the total records by iterating through the recordset
  Recordset2_total=0
  While (Not Recordset2.EOF)
    Recordset2_total = Recordset2_total + 1
    Recordset2.MoveNext
  Wend

  ' reset the cursor to the beginning
  If (Recordset2.CursorType > 0) Then
    Recordset2.MoveFirst
  Else
    Recordset2.Requery
  End If

  ' set the number of rows displayed on this page
  If (Recordset2_numRows < 0 Or Recordset2_numRows > Recordset2_total) Then
    Recordset2_numRows = Recordset2_total
  End If

  ' set the first and last displayed record
  Recordset2_first = 1
  Recordset2_last = Recordset2_first + Recordset2_numRows - 1
  
  If (Recordset2_first > Recordset2_total) Then
    Recordset2_first = Recordset2_total
  End If
  If (Recordset2_last > Recordset2_total) Then
    Recordset2_last = Recordset2_total
  End If

End If
%>

<%
Dim MM_paramName 
%>
<%
' *** Move To Record and Go To Record: declare variables

Dim MM_rs
Dim MM_rsCount
Dim MM_size
Dim MM_uniqueCol
Dim MM_offset
Dim MM_atTotal
Dim MM_paramIsDefined

Dim MM_param
Dim MM_index

Set MM_rs    = Recordset2
MM_rsCount   = Recordset2_total
MM_size      = Recordset2_numRows
MM_uniqueCol = ""
MM_paramName = ""
MM_offset = 0
MM_atTotal = false
MM_paramIsDefined = false
If (MM_paramName <> "") Then
  MM_paramIsDefined = (Request.QueryString(MM_paramName) <> "")
End If
%>
<%
' *** Move To Record: handle 'index' or 'offset' parameter

if (Not MM_paramIsDefined And MM_rsCount <> 0) then

  ' use index parameter if defined, otherwise use offset parameter
  MM_param = Request.QueryString("index")
  If (MM_param = "") Then
    MM_param = Request.QueryString("offset")
  End If
  If (MM_param <> "") Then
    MM_offset = Int(MM_param)
  End If

  ' if we have a record count, check if we are past the end of the recordset
  If (MM_rsCount <> -1) Then
    If (MM_offset >= MM_rsCount Or MM_offset = -1) Then  ' past end or move last
      If ((MM_rsCount Mod MM_size) > 0) Then         ' last page not a full repeat region
        MM_offset = MM_rsCount - (MM_rsCount Mod MM_size)
      Else
        MM_offset = MM_rsCount - MM_size
      End If
    End If
  End If

  ' move the cursor to the selected record
  MM_index = 0
  While ((Not MM_rs.EOF) And (MM_index < MM_offset Or MM_offset = -1))
    MM_rs.MoveNext
    MM_index = MM_index + 1
  Wend
  If (MM_rs.EOF) Then 
    MM_offset = MM_index  ' set MM_offset to the last possible record
  End If

End If
%>
<%
' *** Move To Record: if we dont know the record count, check the display range

If (MM_rsCount = -1) Then

  ' walk to the end of the display range for this page
  MM_index = MM_offset
  While (Not MM_rs.EOF And (MM_size < 0 Or MM_index < MM_offset + MM_size))
    MM_rs.MoveNext
    MM_index = MM_index + 1
  Wend

  ' if we walked off the end of the recordset, set MM_rsCount and MM_size
  If (MM_rs.EOF) Then
    MM_rsCount = MM_index
    If (MM_size < 0 Or MM_size > MM_rsCount) Then
      MM_size = MM_rsCount
    End If
  End If

  ' if we walked off the end, set the offset based on page size
  If (MM_rs.EOF And Not MM_paramIsDefined) Then
    If (MM_offset > MM_rsCount - MM_size Or MM_offset = -1) Then
      If ((MM_rsCount Mod MM_size) > 0) Then
        MM_offset = MM_rsCount - (MM_rsCount Mod MM_size)
      Else
        MM_offset = MM_rsCount - MM_size
      End If
    End If
  End If

  ' reset the cursor to the beginning
  If (MM_rs.CursorType > 0) Then
    MM_rs.MoveFirst
  Else
    MM_rs.Requery
  End If

  ' move the cursor to the selected record
  MM_index = 0
  While (Not MM_rs.EOF And MM_index < MM_offset)
    MM_rs.MoveNext
    MM_index = MM_index + 1
  Wend
End If
%>
<%
' *** Move To Record: update recordset stats

' set the first and last displayed record
Recordset2_first = MM_offset + 1
Recordset2_last  = MM_offset + MM_size

If (MM_rsCount <> -1) Then
  If (Recordset2_first > MM_rsCount) Then
    Recordset2_first = MM_rsCount
  End If
  If (Recordset2_last > MM_rsCount) Then
    Recordset2_last = MM_rsCount
  End If
End If

' set the boolean used by hide region to check if we are on the last record
MM_atTotal = (MM_rsCount <> -1 And MM_offset + MM_size >= MM_rsCount)
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
<%
' *** Move To Record: set the strings for the first, last, next, and previous links

Dim MM_keepMove
Dim MM_moveParam
Dim MM_moveFirst
Dim MM_moveLast
Dim MM_moveNext
Dim MM_movePrev

Dim MM_urlStr
Dim MM_paramList
Dim MM_paramIndex
Dim MM_nextParam

MM_keepMove = MM_keepBoth
MM_moveParam = "index"

' if the page has a repeated region, remove 'offset' from the maintained parameters
If (MM_size > 1) Then
  MM_moveParam = "offset"
  If (MM_keepMove <> "") Then
    MM_paramList = Split(MM_keepMove, "&")
    MM_keepMove = ""
    For MM_paramIndex = 0 To UBound(MM_paramList)
      MM_nextParam = Left(MM_paramList(MM_paramIndex), InStr(MM_paramList(MM_paramIndex),"=") - 1)
      If (StrComp(MM_nextParam,MM_moveParam,1) <> 0) Then
        MM_keepMove = MM_keepMove & "&" & MM_paramList(MM_paramIndex)
      End If
    Next
    If (MM_keepMove <> "") Then
      MM_keepMove = Right(MM_keepMove, Len(MM_keepMove) - 1)
    End If
  End If
End If

' set the strings for the move to links
If (MM_keepMove <> "") Then 
  MM_keepMove = Server.HTMLEncode(MM_keepMove) & "&"
End If

MM_urlStr = Request.ServerVariables("URL") & "?" & MM_keepMove & MM_moveParam & "="

MM_moveFirst = MM_urlStr & "0"
MM_moveLast  = MM_urlStr & "-1"
MM_moveNext  = MM_urlStr & CStr(MM_offset + MM_size)
If (MM_offset - MM_size < 0) Then
  MM_movePrev = MM_urlStr & "0"
Else
  MM_movePrev = MM_urlStr & CStr(MM_offset - MM_size)
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
    <td>
      <table cellpadding=0 cellspacing=0 border=0 width=98% align=center>
        <tr>
	<td width="98" valign="center" height=40><a href="add.asp"><img src="images/postnew.gif" width="72" height="21" border="0" /></a></td>
          <td width="98"  valign="center"  ><a href="#point"><img src="images/newreply.gif" width="72" height="21" border="0" alt="回复主题"></a></td>
	<td align=right valign=middle><span class="STYLE2">您是本帖的第<%=(Recordset1.Fields.Item("num_hits").Value)%>个阅读者</span>&nbsp; &nbsp; 
      　 </td>
	</tr>
</table>
      <table border="0" width="98%" cellspacing="1" cellpadding="3" bgcolor="#205E7B" align="center">
        <tr bgcolor="#205E7B"> 
          <td colspan="2" height="25" class="td1">&nbsp;&nbsp;<b><font color="#FFFFFF"> 
            <img src="<%=(Recordset1.Fields.Item("main_important").Value)%>" vspace="2" />标题：<%=(Recordset1.Fields.Item("main_subject").Value)%></font></b></td>
        </tr>
        <tr> 
          <td bgcolor="#DAEDF5" align="center" valign="top" width="20%"> 
            <table width="98%" border="0" cellspacing="0" cellpadding="0" align="center">
              <tr> 
                <td><p align="center"><img border="0" src="images/bl.gif" /><span class="STYLE2"><%=(Recordset1.Fields.Item("main_name").Value)%></span><br />
                </p>
                <p align="center"><img src=<%=(Recordset1.Fields.Item("main_face").Value)%> width="100" height="100" alt="" /></p>
                </td>
              </tr>
            </table>
          </td>
          <td width="80%" align="center" valign="top" bgcolor="#FFFFFF"> 
            <table width="98%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td width="56%"><a href=<%=(Recordset1.Fields.Item("main_url").Value)%>><img src="images/homepage.gif" width="47" height="18" border="0" /></a> <a href=mailto:<%=(Recordset1.Fields.Item("main_email").Value)%>><img src="images/email.gif" width="45" height="18" border="0" /></a></td>
                <td width="44%" align="right" class="STYLE2"><%=(Recordset1.Fields.Item("main_time").Value)%></td>
              </tr>
              <tr> 
                <td height="1" bgcolor="#999999" colspan="2"></td>
              </tr>
            </table>
            <table width="80%" border="0" cellspacing="0" cellpadding="10" align="center"style="TABLE-LAYOUT: fixed">
              <tr> 
                <td width="533" valign="top" style="LEFT: 20px; WIDTH: 100%; WORD-WRAP: break-word"> 
                  <p class="STYLE2" style="line-height: 150%">
                  <%=(Recordset1.Fields.Item("main_content").Value)%></p>
                <p style="line-height: 150%">                </p></td>
              </tr>
            </table>
          </td>
        </tr>
        <tr> 
          <td bgcolor="#92C8E2" align="center" valign="top" width="20%">&nbsp;</td>
          <td bgcolor="#92C8E2" width="80%">&nbsp;</td>
        </tr>
        <% 
While ((Repeat1__numRows <> 0) AND (NOT Recordset2.EOF)) 
%> <tr>
         
<td bgcolor="#DAEDF5" align="center" valign="top" width="20%"> 
            <table width="98%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td align="center">&nbsp;<span class="STYLE2"><img border="0" src="images/bl.gif"><%=(Recordset2.Fields.Item("re_name").Value)%></span></td>
              </tr>
            </table>
            <img src=<%=(Recordset2.Fields.Item("re_face").Value)%>  width="100" height="100" alt="" /><br>
          </td>
          <td width="80%" align="right" valign="top" bgcolor="#FFFFFF"> 
            <table width="98%" border="0" cellspacing="0" cellpadding="0" align="center">
              <tr> 
                <td width="59%"><span style="line-height: 150%"><b><font color="#FF0000"><a href=<%=(Recordset2.Fields.Item("re_url").Value)%>><img src="images/homepage.gif" width="47" height="18" border="0" /></a> <a href="mailto:<%=(Recordset2.Fields.Item("re_email").Value)%>"><img src="images/email.gif" width="45" height="18" border="0" /></a></font></b></span></td>
                <td width="41%" align="right" class="STYLE2"><%=(Recordset2.Fields.Item("re_time").Value)%></td>
              </tr>
              <tr> 
                <td height="1" bgcolor="#999999" colspan="2"></td>
              </tr>
            </table>
            <table width="80%" border="0" cellspacing="0" cellpadding="10" align="center"style="TABLE-LAYOUT: fixed">
              <tr> 
                <td width="533" valign="top" style="LEFT: 20px; WIDTH: 100%; WORD-WRAP: break-word"> 
                  <p style="line-height: 150%"><b><font color="#FF0000">标题：</font></b><span class="STYLE2"><%=(Recordset2.Fields.Item("re_subject").Value)%><br />
                      <%=(Recordset2.Fields.Item("re_content").Value)%></span><br>
                </p>                </td>
              </tr>
            </table>
          </td>
        </tr>
        <tr>
          <td bgcolor="#92C8E2" align="center" valign="top" width="20%">&nbsp;</td>
          <td bgcolor="#92C8E2" width="80%" align=right>&nbsp;</td>
</tr>
        
          <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  Recordset2.MoveNext()
Wend
%>
      </table>

      <table width="98%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td width="20%">共有<%=(Recordset2_total)%>条记录</td>
          <td width="80%" align="right" height="30">&nbsp;
            <table border="0">
              <tr>
              <td><% If MM_offset <> 0 Then %>
                      <a href="<%=MM_moveFirst%>">第一页</a>
                      <% End If ' end MM_offset <> 0 %>
                </td>
              <td><% If MM_offset <> 0 Then %>
                      <a href="<%=MM_movePrev%>">前一页</a>
                      <% End If ' end MM_offset <> 0 %>
                </td>
              <td><% If Not MM_atTotal Then %>
                      <a href="<%=MM_moveNext%>">下一页</a>
                      <% End If ' end Not MM_atTotal %>
                </td>
              <td><% If Not MM_atTotal Then %>
                      <a href="<%=MM_moveLast%>">最后一页</a>
                      <% End If ' end Not MM_atTotal %>
                </td>
              </tr>
            </table></td>
        </tr>
      </table>
      <form METHOD="POST"  name="form1" action="<%=MM_editAction%>">
      <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" bordercolorlight="#999999" bordercolordark="#FFFFFF" bgcolor="#205E7B" class="STYLE2">
		<tr> 
            <td colspan="3" height="25" class="td1">　　<A name=point class="a1">回复主题：<%=(Recordset1.Fields.Item("main_subject").Value)%></A><input type="hidden" name="subject" value="<%=(Recordset1.Fields.Item("main_subject").Value)%>" />
            <input type="hidden" name="id" value="<%=(Recordset1.Fields.Item("main_id").Value)%>" /></td>
        </tr>
            <tr> 
            <td width="21%" align="center" bgcolor="#92C8E2">发&nbsp; 表&nbsp; 人：</td>
              <td colspan="2" bgcolor="#FFFFFF"> 
                    <input type="text" name=name size="20" value="" >
                   （ID限制 <b>5</b> 汉字以内）<font color="#FF0000">*</font> </td></tr>
            <tr> 
            <td width="21%" align="center" bgcolor="#92C8E2">电 子 邮 件：</td>
              
            <td width="19%" bgcolor="#92C8E2"> 
              <input type="text" name="email" size="20" >            </td>
              
            <td width="60%" bgcolor="#92C8E2">主 页 地 址： 
              <input type="text" name="url" size="30"  value="http://">            </td>
            </tr>
            <tr> 
            <td width="21%" align="center" bgcolor="#DAEDF5"><img id=idface src="images/01.gif" alt=个人形象代表><br> 
              <select name="face" size=1 onChange="document.images['idface'].src=options[selectedIndex].value;" >
                <option selected value="images/01.gif">用户头像-01 
                <option selected value="images/02.gif">用户头像-02 
                <option selected value="images/03.gif">用户头像-03 
                <option selected value="images/04.gif">用户头像-04 
                <option selected value="images/05.gif">用户头像-05 
                <option selected value="images/06.gif">用户头像-06 
                <option selected value="images/07.gif">用户头像-07 
                <option selected value="images/08.gif">用户头像-08 
                <option selected value="images/09.gif">用户头像-09 
                <option selected value="images/10.gif">用户头像-10 
                <option selected value="images/11.gif">用户头像-11 
                <option selected value="images/12.gif">用户头像-12 
                <option selected value="images/13.gif">用户头像-13 
                <option selected value="images/14.gif">用户头像-14 
                <option selected value="images/15.gif">用户头像-15 
                <option selected value="images/16.gif">用户头像-16 
                <option selected value="images/17.gif">用户头像-17 
                <option selected value="images/18.gif">用户头像-18 
                <option selected value="images/19.gif">用户头像-19 
                <option selected value="images/20.gif">用户头像-20 
              </select>              </td>
              
            <td valign="top" colspan="2" bgcolor="#DAEDF5">
              <label>
                <textarea name="textarea" id="textarea" cols="100" rows="10"></textarea>
              </label>            </td>
            </tr>
          <tr> 
            <td colspan="3"> 
              <p align="center">
<input type="submit" value=" 提交 " name="B1" >
 &nbsp;
 <input type="reset" value=" 清除 " name="B2" >
            </td>
          </tr>
      </table>
      <input type="hidden" name="MM_insert" value="form1" />
      </form>
     
      <br>
    </td>
  </tr>
</table>
<table width="98%" border="0" cellspacing="0" cellpadding="0" align="center" bgcolor="#000000">
  <tr>
    <td align="center" height="30" class=td1>
    </td>
  </tr>
</table>
</body>
</html><%
Recordset1.Close()
Set Recordset1 = Nothing
%>
<%
Recordset2.Close()
Set Recordset2 = Nothing
%>
