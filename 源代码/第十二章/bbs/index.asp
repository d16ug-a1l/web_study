<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="Connections/connection.asp" -->
<%
Dim Recordset1
Dim Recordset1_cmd
Dim Recordset1_numRows

Set Recordset1_cmd = Server.CreateObject ("ADODB.Command")
Recordset1_cmd.ActiveConnection = MM_connection_STRING
Recordset1_cmd.CommandText = "SELECT * FROM postMain ORDER BY main_time DESC" 
Recordset1_cmd.Prepared = true

Set Recordset1 = Recordset1_cmd.Execute
Recordset1_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = 2
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

Set MM_rs    = Recordset1
MM_rsCount   = Recordset1_total
MM_size      = Recordset1_numRows
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
Recordset1_first = MM_offset + 1
Recordset1_last  = MM_offset + MM_size

If (MM_rsCount <> -1) Then
  If (Recordset1_first > MM_rsCount) Then
    Recordset1_first = MM_rsCount
  End If
  If (Recordset1_last > MM_rsCount) Then
    Recordset1_last = MM_rsCount
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
.STYLE1 {
	color: #FFFFFF;
	font-weight: bold;
}
.STYLE2 {
	font-size: 14px;
	color: #000000;
}
.bg {
	background-repeat: no-repeat;
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
-->
</style>
<title>论坛首页</title><body>
<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0" >
  <tr> <td></td>
                <td  height=37 align=right valign="bottom" background="images/top.jpg" bgcolor="#6EB7DA" class="bg"><br>
                  <br>
                  <a href="index.asp">论坛首页</a>&nbsp;&nbsp;&nbsp;&nbsp;<a href="login.asp">管理登陆</a>
    &nbsp;</td>
  </tr>
</table>
<table width="98%" border="0" cellspacing="0" cellpadding="0" align="center" >
  <tr>
    <td>
      <table border="0" width="98%" cellspacing="0" cellpadding="0" align="center">
  <tr> 
    <td width="98"  valign="center" height=40><a href="add.asp"><img src="images/postnew.gif" width="72" height="21" border="0"></a></td>
    
        <td  valign="center" >&nbsp;</td>
        </tr>
</table>
      <table width="98%" border="0" cellspacing="1" cellpadding="0" bgcolor="#205E7B" align="center">
        <tr bgcolor="#f0f0f0"> 
          <td width="80%">
      <table border="0" width="98%" cellspacing="1" cellpadding="0"  bgcolor="#205E7B" align="center" >
        <tr> 
          <td width="350" height="25" align="center"><b class="STYLE1"><font color="#FFFFFF">帖 子 主 题</font></b></td>
          <td align="center" width="209"><b class="STYLE1"><font color="#FFFFFF">作 者</font></b></td>
          <td align="center" width="147"><span class="STYLE1">点击次数</span></td>
          <td align="center" width="196" ><b class="STYLE1"><font color="#FFFFFF">创建时间</font></b></td>
        </tr>
        <% 
While ((Repeat1__numRows <> 0) AND (NOT Recordset1.EOF)) 
%>
          <tr bgcolor="#FFFFFF"> 
              <td height="30" align="center"> 
                  <div align="left"><img src=<%=(Recordset1.Fields.Item("main_important").Value)%> vspace="2"><span class="STYLE2">&nbsp;&nbsp;</span><span class="STYLE2"><a href=show.asp?main_id=<%=(Recordset1.Fields.Item("main_id").Value)%>><%=(Recordset1.Fields.Item("main_subject").Value)%></a></span></div></td>
            <td width="209" align="center" bgcolor="#92C8E2" class="STYLE2"><%=(Recordset1.Fields.Item("main_name").Value)%></td>
            <td width="147" align="center" bgcolor="#DAEDF5" class="STYLE2"><%=(Recordset1.Fields.Item("num_hits").Value)%></td>
            <td bgcolor="#92C8E2" width="196">&nbsp;&nbsp;<span class="STYLE2"><%=(Recordset1.Fields.Item("main_time").Value)%></span></td>
          </tr>
          <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  Recordset1.MoveNext()
Wend
%>
            </table>
      <table width="98%" border="0" align="center" cellspacing="0" cellpadding="3" bgcolor="#205E7B">
        <form method=post action=index.asp>
          <tr> 
            <td width="52%" class=td1><% If Recordset1.EOF And Recordset1.BOF Then %>
                暂时没有帖子！
  <% End If ' end Recordset1.EOF And Recordset1.BOF %>
</td>
            <td width="48%" align="right" class=td1><table border="0">
                <tr>
                  <td><% If MM_offset <> 0 Then %>
                        <span ><a href="<%=MM_moveFirst%>">第一页</a></span><a href="<%=MM_moveFirst%>"><br>
                        </a>
                        <% End If ' end MM_offset <> 0 %>
                  </td>
                  <td><% If MM_offset <> 0 Then %>
                        <a href="<%=MM_movePrev%>" >前一页</a>
                        <% End If ' end MM_offset <> 0 %>
                  </td>
                  <td><% If Not MM_atTotal Then %>
                        <a href="<%=MM_moveNext%>" >下一页</a>
                        <% End If ' end Not MM_atTotal %>
                  </td>
                  <td><% If Not MM_atTotal Then %>
                        <a href="<%=MM_moveLast%>" >最后一页</a>
                        <% End If ' end Not MM_atTotal %>
                  </td>
                </tr>
              </table></td>
          </tr>
        </form>
      </table>
	  <table width="98%" border="0" cellspacing="0" cellpadding="0" align="center">
        <tr>
   <form name="searchtitle" method="POST" action="search.asp" target="_blank">
            <td align="left">快速搜索：
              <input name="keyword" type="text"  size="16">
              <input type="submit" name="Submit" value="搜 索">
            </td>
   </form>
        </tr>
      </table>
      
 
      <table cellspacing=1 cellpadding=3 width="98%" bgcolor="#205E7B" align="center">
        <tr> 
    <td ><font color="#FFFFFF">　-=&gt; <b>BBS图例</b></font></td>
  </tr>
  <tr> 
    <td colspan=2 bgcolor="#FFFFFF"> 
      <table cellspacing=4 cellpadding=0 width="92%" align=center border=0 bgcolor="#FFFFFF">
        <tr> 
          <td><img src="images/putong.gif"> 普通帖子</td>
          <td><img src="images/hot.gif"> 热门帖子</td>
          <td><img src="images/isbest.gif"> 精华帖子 </td>
        </tr>
      </table>
    </td>
  </tr>
</table>
      <br>
    </td>
  </tr>
</table>
<table width="98%" border="0" cellspacing="0" cellpadding="0" align="center" bgcolor="#20eEff">
  <tr> 
    <td align="center" height="30"class=STYLE2 width="96%">CopyRight 21世纪网络</td>
    
  </tr>
</table>
</body>
</html>
<%
Recordset1.Close()
Set Recordset1 = Nothing
%>