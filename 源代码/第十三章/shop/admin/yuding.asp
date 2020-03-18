<%@LANGUAGE="VBSCRIPT"%>
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
<html>
<head>
<title>Freedom</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type="text/css">
<!--
.STYLE1 {
	color: #0000FF;
	font-size: 14px;
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
<script src="../Scripts/AC_RunActiveContent.js" type="text/javascript"></script>
</head>

<body bgcolor="#FFFFFF" text="#333333" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" >
<table width="100%" border="0" cellspacing="0" cellpadding="0" height="100%">
  <tr>
    <td height="361" valign="top"> 
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td background="image/bg_top2.gif" height="67" valign="top"> 
            <table width="536" border="0" cellspacing="0" cellpadding="0" background="">
              <tr> 
                <td width="190"><img src="image/transparent.gif" width="10" height="20"></td>
                <td width="346" align="right"><a href="../admin/login.asp"></a></td>
              </tr>
              <tr> 
                <td width="190"><img src="image/transparent.gif" width="10" height="22"></td>
                <td width="346">　</td>
              </tr>
              <tr>
                <td width="190">
                  <table border="0" cellspacing="0" cellpadding="0">
                    <tr>
                      <td><img src="images/obj_title.gif" width="190" height="28"></td>
                    </tr>
                    <tr>
                      <td><img src="image/transparent.gif" width="10" height="1"></td>
                    </tr>
                  </table>                </td>
                <td width="346" valign="bottom"><table border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td><a href="../index.html"><img src="image/bt_01_off.gif" width="86" height="18" border="0" name="Image5"></a></td>
                    <td><a href="adminchakan.asp"><img src="images/bt8.gif" width="86" height="18" border="0" name="Image1"></a></td>
                    <td><a href="userchakan.asp"><img src="images/bt2.gif" width="86" height="18" border="0" name="Image2"></a></td>
                    <td><a href="shopchakan.asp"><img src="images/bt5.gif" width="86" height="18" border="0" name="Image3"></a></td>
                    <td><a href="yuding.asp"><img src="images/bt6.gif" width="86" height="18" border="0" name="Image4"></a></td>
                    <td><img src="image/obj_bt.gif" width="9" height="18"></td>
                  </tr>
                </table></td>
              </tr>
              <tr> 
                <td width="190"><img src="image/transparent.gif" width="10" height="25"></td>
                <td width="346">　</td>
              </tr>
            </table>          </td>
        </tr>
        <tr>
          <td height="5" valign="top"><img src="image/transparent.gif" width="20" height="10"></td>
        </tr>
      </table>
      <table width="640" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="16"><img src="image/transparent.gif" width="50" height="20"></td>
          <td width="618"><table width="668" border="0" cellpadding="0" cellspacing="0">
              <tr>
                <td width="668" align="left" valign="middle"><table width="607" border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td width="143" height="28" class="STYLE1"><div align="center"><a href="shoptianjia.asp"></a><a href="yuding.asp">查看购买成功信息</a></div></td>
                    <td width="183" class="STYLE1"><div align="center"><a href="yudingno.asp">查看购买未成功信息</a></div></td>
                    <td width="281" class="STYLE1"><div align="center"><a href="yudingxg.asp">修改订单信息</a></div></td>
                  </tr>
                  <tr>
                    <td colspan="3"><p>&nbsp;</p>
                      <p>&nbsp;</p></td>
                  </tr>
                </table>
                  <img src="image/line.gif" width="590" height="11">
                  <table align="center" cellpadding="0" cellspacing="1"bordercolor="#99BB99" bgcolor="#0000FF" style=" border-collapse: collapse">
                 
                      <tr>
                        <td align="right" bgcolor="#FFFFFF"><p align="center"  class="greenb">用户名</td>
                        <td bgcolor="#FFFFFF" ><div align="center"><span class="greenb">用户真实姓名</span></div></td>
                        <td bgcolor="#FFFFFF" ><div align="center"><span class="greenb">用户地址</span><span class="greenb"></span></div></td>
                        <td bgcolor="#FFFFFF" ><div align="center"><span class="greenb">用户</span><span class="greenb">电话</span></div></td>
                        <td bgcolor="#FFFFFF" ><div align="center"><span class="greenb">商品名称</span></div></td>
                        <td bgcolor="#FFFFFF" ><div align="center">汇款方式</div></td>
                      </tr>
                      <tr>
                        <td align="center" bgcolor="#FFFFFF">&nbsp;</td>
                        <td align="center" bgcolor="#FFFFFF">&nbsp;</td>
                        <td align="center" bgcolor="#FFFFFF"><p>&nbsp;</p></td>
                        <td align="center" bgcolor="#FFFFFF">&nbsp;</td>
                        <td align="center" bgcolor="#FFFFFF">&nbsp;</td>
                        <td align="center" bgcolor="#FFFFFF">&nbsp;</td>
                        <td align="center"><span class="STYLE1"><a href="zhuce.asp"></a></span></td>
                      </tr>
                  </table>
                  <br>
第一页&nbsp;&nbsp;&nbsp;前一页&nbsp;&nbsp;&nbsp;下一页&nbsp;&nbsp;&nbsp;最后一页&nbsp;&nbsp;&nbsp;共条记录<br>
                <br></td>
              </tr>
          </table></td>
        </tr>
      </table> </td>
  </tr>
  <tr>
    <td valign="bottom"><table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td background="image/bg_bottom.gif" align="right">
            <table border="0" cellspacing="0" cellpadding="0" background="">
              <tr> 
                <td height="14"><img src="image/transparent.gif" width="210" height="20"></td>
                <td rowspan="2"><img src="image/transparent.gif" width="3" height="40"></td>
              </tr>
            </table>          </td>
        </tr>
    </table>    </td>
  </tr>
</table>
</body>
</html>