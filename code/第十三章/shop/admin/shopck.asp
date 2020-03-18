<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../Connections/connection.asp" -->
<%
Dim Recordset1__MMColParam
Recordset1__MMColParam = "�ֻ���"
If (Request.QueryString("leibie") <> "") Then 
  Recordset1__MMColParam = Request.QueryString("leibie")
End If
%>
<%
Dim Recordset1
Dim Recordset1_cmd
Dim Recordset1_numRows

Set Recordset1_cmd = Server.CreateObject ("ADODB.Command")
Recordset1_cmd.ActiveConnection = MM_connection_STRING
Recordset1_cmd.CommandText = "SELECT * FROM shop WHERE sp_leibie = ?" 
Recordset1_cmd.Prepared = true
Recordset1_cmd.Parameters.Append Recordset1_cmd.CreateParameter("param1", 200, 1, 50, Recordset1__MMColParam) ' adVarChar

Set Recordset1 = Recordset1_cmd.Execute
Recordset1_numRows = 0
%>
<%
Dim Recordset2
Dim Recordset2_cmd
Dim Recordset2_numRows

Set Recordset2_cmd = Server.CreateObject ("ADODB.Command")
Recordset2_cmd.ActiveConnection = MM_connection_STRING
Recordset2_cmd.CommandText = "SELECT * FROM fenlei" 
Recordset2_cmd.Prepared = true

Set Recordset2 = Recordset2_cmd.Execute
Recordset2_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = 5
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
<html>
<head>
<title>���Ϲ�����վ</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">

<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
}
a {
	font-family: ����;
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
</style>
<script type="text/javascript">
<!--
function MM_jumpMenu(targ,selObj,restore){ //v3.0
  eval(targ+".location='"+selObj.options[selObj.selectedIndex].value+"'");
  if (restore) selObj.selectedIndex=0;
}
//-->
</script>
</head>

<body>
<table width="100%" height="606" border="0" cellpadding="0" cellspacing="0" background="image/bg_top2.gif">
  <tr>
    <td height="89" align="left" valign="top" background="../image/bg_top2.gif">
     <table width="644" height="75" border="0" cellpadding="0" cellspacing="0">
       <tr>
         <td height="20" colspan="6">&nbsp;</td>
         <td width="104" colspan="2"><div align="right"><a href="../index.asp"><img src="images/con_contact.gif" border="0"></a></div></td>
       </tr>
        <tr>
           <td width="190" height="28"><img src="../image/obj_title.gif" width="190" height="28"></td>
          <td colspan="7">&nbsp;</td>
        </tr>
       <tr>
          <td>&nbsp;</td>
          <td width="86"><a href="index.asp"><img src="images/bt9.gif" width="86" height="18" border="0"></a></td>
          <td width="86"><a href="che.asp"><img src="images/bt1.gif" width="86" height="18" border="0"></a></td>
          <td width="83"><a href="index.asp"><img src="images/bt3.gif" width="86" height="18" border="0"></a></td>
          <td width="55"><img src="images/bt.gif" width="86" height="18"></td>
          <td><img src="images/bt14.gif" width="86" height="18"></td>
          <td><a href="shopck.asp"><img src="images/bt5.gif" width="86" height="18" border="0"></a></td>
          <td><a href="dingdan.asp"><img src="images/bt7.gif" width="86" height="18" border="0"></a></td>
       </tr>
    </table>
   </td>
  </tr>
  <tr>
    <td height="154" align="left" valign="top">
    <table width="100%" height="476" border="0" cellpadding="0" cellspacing="0" bgcolor="#E7E7E7">
      <tr>
        <td height="27" align="right"><form name="form" id="form">
          <div align="center">
            <select name="jumpMenu" id="jumpMenu" onChange="MM_jumpMenu('parent',this,0)">
              <option value="ѡ�����">ѡ�����</option>
              <%
While (NOT Recordset2.EOF)
%>
              <option value="shopck.asp?leibie=<%=(Recordset2.Fields.Item("fenlei").Value)%>"><%=(Recordset2.Fields.Item("fenlei").Value)%></option>
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
          </div>
        </form>
          <p>&nbsp;</p>
          
          <% If Not Recordset1.EOF Or Not Recordset1.BOF Then %>
            <table width="782" align="center" cellpadding="0" cellspacing="1"bordercolor="#99BB99" bgcolor="#0000FF" style=" border-collapse: collapse">
              <tr>
                <td width=158 align="right" bgcolor="#FFFFFF"><p align="center"  class="greenb">��Ʒ����</td>
                  <td width="158" bgcolor="#FFFFFF" ><div align="center"><span class="greenb">��Ʒ</span><span class="greenb">�۸�</span></div></td>
                  <td width="227" bgcolor="#FFFFFF" ><div align="center">��Ʒ���</div></td>
                  <td width="154" bgcolor="#FFFFFF" >��ƷƷ��</td>
                  <td width="39" align="center" bgcolor="#FFFFFF" >ɾ��</td>
                  <td width="37" bgcolor="#FFFFFF" >�޸�</td>
              </tr>
              <% 
While ((Repeat1__numRows <> 0) AND (NOT Recordset1.EOF)) 
%>
                <tr>
                  <td height="50" align="center" bgcolor="#FFFFFF"><%=(Recordset1.Fields.Item("sp_title").Value)%></td>
                  <td align="center" bgcolor="#FFFFFF"><%=(Recordset1.Fields.Item("sp_price").Value)%></td>
                  <td align="center" bgcolor="#FFFFFF"><%=(Recordset1.Fields.Item("sp_shuoming").Value)%></td>
                  <td align="center" bgcolor="#FFFFFF"><%=(Recordset1.Fields.Item("sp_pinpai").Value)%></td>
                  <td align="center" bgcolor="#FFFFFF"><a href="delshop.asp?id=<%=(Recordset1.Fields.Item("id").Value)%>">ɾ��</a></td>
                  <td align="center" bgcolor="#FFFFFF"><a href="updateshop.asp?id=<%=(Recordset1.Fields.Item("id").Value)%>">�޸�</a></td>
                </tr>
                <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  Recordset1.MoveNext()
Wend
%>
            </table>
            <% End If ' end Not Recordset1.EOF Or NOT Recordset1.BOF %>
<div align="right">
            <table border="0" align="left">
              <tr>
                <td><% If MM_offset <> 0 Then %>
                  <a href="<%=MM_moveFirst%>">��һҳ</a>
                  <% End If ' end MM_offset <> 0 %>                  </td>
                  <td><% If MM_offset <> 0 Then %>
                    <a href="<%=MM_movePrev%>">ǰһҳ</a>
                    <% End If ' end MM_offset <> 0 %>                  </td>
                  <td><% If Not MM_atTotal Then %>
                    <a href="<%=MM_moveNext%>">��һҳ</a>
                    <% End If ' end Not MM_atTotal %>                  </td>
                  <td><% If Not MM_atTotal Then %>
                    <a href="<%=MM_moveLast%>">���һҳ</a>
                    <% End If ' end Not MM_atTotal %>                  </td>
                </tr>
            </table>
            <p>&nbsp;</p>
            <% If Recordset1.EOF And Recordset1.BOF Then %>
              <p align="center">û�и÷������Ʒ.</p>
              <% End If ' end Recordset1.EOF And Recordset1.BOF %>
</div></td>
        </tr>
      <tr>
        <td height="449" align="left" valign="top">&nbsp;</td>
        </tr>
    </table></td>
  </tr>
  <tr>
    <td height="38" background="../image/bg_bottom.gif"><img src="../image/transparent.gif" height="1"></td>
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
