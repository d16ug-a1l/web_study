<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="Connections/connection.asp" -->
<%
Dim Recordset1__MMColParam
Recordset1__MMColParam = "1"
If (Session("orderid") <> "") Then 
  Recordset1__MMColParam = Session("orderid")
End If
%>
<%
Dim Recordset1
Dim Recordset1_cmd
Dim Recordset1_numRows

Set Recordset1_cmd = Server.CreateObject ("ADODB.Command")
Recordset1_cmd.ActiveConnection = MM_connection_STRING
Recordset1_cmd.CommandText = "SELECT orderid, sp_id, num, sp_title, sp_price, sp_num, dingdan.id FROM dingdan, shop WHERE orderid = ? AND shop.id=dingdan.sp_id" 
Recordset1_cmd.Prepared = true
Recordset1_cmd.Parameters.Append Recordset1_cmd.CreateParameter("param1", 5, 1, -1, Recordset1__MMColParam) ' adDouble

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
<title>网上购物网站</title>
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
<table width="100%" height="606" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td height="89" align="left" valign="top" background="image/bg_top2.gif">
     <table width="644" height="75" border="0" cellpadding="0" cellspacing="0">
       <tr>
         <td height="20" colspan="6">&nbsp;</td>
         <td width="104"><div align="right"><a href="admin/login.asp"><img src="image/con_contact.gif" border="0"></a></div></td>
       </tr>
        <tr>
           <td width="190" height="28"><img src="image/obj_title.gif" width="190" height="28"></td>
           <td colspan="6">&nbsp;</td>
        </tr>
       <tr>
          <td>&nbsp;</td>
          <td width="86"><a href="index.asp"><img src="image/bt_01_off.gif" width="86" height="18" border="0"></a></td>
          <td width="86"><a href="index.asp"><img src="image/bt_03_off.gif" width="86" height="18" border="0"></a></td>
          <td width="83"><a href="sousuo.asp"><img src="image/bt_04_off.gif" width="86" height="18" border="0"></a></td>
          <td width="55"><a href="dingdan.asp"><img src="image/bt_08_off.gif" width="86" height="18" border="0"></a></td>
          <td width="40"><a href="che.asp"><img src="image/bt_09_off.gif" width="86" height="18" border="0"></a></td>
          <td><a href="mailto:fuping@sohu.co"><img src="image/bt_05_off.gif" width="86" height="18" border="0"></a></td>
      </tr>
    </table>
   </td>
  </tr>
  <tr>
    <td height="154" align="left" valign="top" bgcolor="#E7E7E7">
      <table width="900" height="449" border="0" cellpadding="0" cellspacing="0" bgcolor="#E7E7E7">
      <tr>
        <td height="449" align="left" valign="top"><% If Not Recordset1.EOF Or Not Recordset1.BOF Then %>
            <form action="chexiugai.asp" name="form1" >
              <table width="90%" border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td width="19%"><div align="center">商品名称</div></td>
                  <td width="20%"><div align="center">商品价格</div></td>
                  <td width="17%"><div align="center">商品数量</div></td>
                  <td width="17%"><div align="center">删除</div></td>
                </tr>
                <% 
While ((Repeat1__numRows <> 0) AND (NOT Recordset1.EOF))
%>
                  <tr>
                    <td width="19%"><div align="center"><%=(Recordset1.Fields.Item("sp_title").Value)%></div></td>
                    <td width="20%"><label>
                        <div align="center"><%=(Recordset1.Fields.Item("sp_price").Value)%></div>
                      </label></td>
                    <td width="17%"><label>
                        <div align="center">
                          <input name="<%=(Recordset1.Fields.Item("id").Value)%>" type="text" value="<%=(Recordset1.Fields.Item("num").Value)%>">
                          </label>
                      </div></td>
                    <td width="17%"><div align="center">删除</div></td>
                   <% s=s+(Recordset1.Fields.Item("sp_price").Value)*(Recordset1.Fields.Item("num").Value) %>
                  </tr>
                  <% 
	  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  Recordset1.MoveNext()
Wend
%>
              </table>
              <p>总价格为：<%=s%><% Session("price")=s %></p>
              <p>
                <label>
                <input type="submit" name="button" id="button" value="修改订单">
                </label>
              <a href="javascript:history.go(-1)">继续购物</a>
               <a href="goumai.asp">购买商品</a></p>
            </form>
            <% End If ' end Not Recordset1.EOF Or NOT Recordset1.BOF %>
          <% If Recordset1.EOF And Recordset1.BOF Then %>
            <p>您还没有购物！</p>
            <% End If ' end Recordset1.EOF And Recordset1.BOF %></td>
        </tr>
    </table></td>
  </tr>
  <tr>
    <td height="38" background="image/bg_bottom.gif"><img src="image/transparent.gif" height="40"></td>
  </tr>
</table>
</body>
</html>
<%
Recordset1.Close()
Set Recordset1 = Nothing
%>
