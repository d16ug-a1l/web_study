<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../Connections/connnection.asp" -->
<%
' *** Edit Operations: declare variables

Dim MM_editAction
Dim MM_abortEdit
Dim MM_editQuery
Dim MM_editCmd

Dim MM_editConnection
Dim MM_editTable
Dim MM_editRedirectUrl
Dim MM_editColumn
Dim MM_recordId

Dim MM_fieldsStr
Dim MM_columnsStr
Dim MM_fields
Dim MM_columns
Dim MM_typeArray
Dim MM_formVal
Dim MM_delim
Dim MM_altVal
Dim MM_emptyVal
Dim MM_i

MM_editAction = CStr(Request.ServerVariables("SCRIPT_NAME"))
If (Request.QueryString <> "") Then
  MM_editAction = MM_editAction & "?" & Server.HTMLEncode(Request.QueryString)
End If

' boolean to abort record edit
MM_abortEdit = false

' query string to execute
MM_editQuery = ""
%>
<%
' *** Update Record: set variables

If (CStr(Request("MM_update")) = "form1" And CStr(Request("MM_recordId")) <> "") Then

  MM_editConnection = MM_connnection_STRING
  MM_editTable = "yuding"
  MM_editColumn = "id"
  MM_recordId = "" + Request.Form("MM_recordId") + ""
  MM_editRedirectUrl = "yudingxg.asp"
  MM_fieldsStr  = "sp_id|value|sp_title|value|yd_personid|value|yd_name|value|phone|value|address|value|fangshi|value|time|value|checkbox|value"
  MM_columnsStr = "sp_id|',none,''|sp_title|',none,''|yd_personid|',none,''|yd_name|',none,''|yd_phone|',none,''|yd_address|',none,''|fanshi|',none,''|yd_time|',none,NULL|yd_ifsuccess|none,1,0"

  ' create the MM_fields and MM_columns arrays
  MM_fields = Split(MM_fieldsStr, "|")
  MM_columns = Split(MM_columnsStr, "|")
  
  ' set the form values
  For MM_i = LBound(MM_fields) To UBound(MM_fields) Step 2
    MM_fields(MM_i+1) = CStr(Request.Form(MM_fields(MM_i)))
  Next

  ' append the query string to the redirect URL
  If (MM_editRedirectUrl <> "" And Request.QueryString <> "") Then
    If (InStr(1, MM_editRedirectUrl, "?", vbTextCompare) = 0 And Request.QueryString <> "") Then
      MM_editRedirectUrl = MM_editRedirectUrl & "?" & Request.QueryString
    Else
      MM_editRedirectUrl = MM_editRedirectUrl & "&" & Request.QueryString
    End If
  End If

End If
%>
<%
' *** Update Record: construct a sql update statement and execute it

If (CStr(Request("MM_update")) <> "" And CStr(Request("MM_recordId")) <> "") Then

  ' create the sql update statement
  MM_editQuery = "update " & MM_editTable & " set "
  For MM_i = LBound(MM_fields) To UBound(MM_fields) Step 2
    MM_formVal = MM_fields(MM_i+1)
    MM_typeArray = Split(MM_columns(MM_i+1),",")
    MM_delim = MM_typeArray(0)
    If (MM_delim = "none") Then MM_delim = ""
    MM_altVal = MM_typeArray(1)
    If (MM_altVal = "none") Then MM_altVal = ""
    MM_emptyVal = MM_typeArray(2)
    If (MM_emptyVal = "none") Then MM_emptyVal = ""
    If (MM_formVal = "") Then
      MM_formVal = MM_emptyVal
    Else
      If (MM_altVal <> "") Then
        MM_formVal = MM_altVal
      ElseIf (MM_delim = "'") Then  ' escape quotes
        MM_formVal = "'" & Replace(MM_formVal,"'","''") & "'"
      Else
        MM_formVal = MM_delim + MM_formVal + MM_delim
      End If
    End If
    If (MM_i <> LBound(MM_fields)) Then
      MM_editQuery = MM_editQuery & ","
    End If
    MM_editQuery = MM_editQuery & MM_columns(MM_i) & " = " & MM_formVal
  Next
  MM_editQuery = MM_editQuery & " where " & MM_editColumn & " = " & MM_recordId

  If (Not MM_abortEdit) Then
    ' execute the update
    Set MM_editCmd = Server.CreateObject("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_editConnection
    MM_editCmd.CommandText = MM_editQuery
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close

    If (MM_editRedirectUrl <> "") Then
      Response.Redirect(MM_editRedirectUrl)
    End If
  End If

End If
%>
<%
Dim Recordset1__MMColParam
Recordset1__MMColParam = "1"
If (Request.QueryString("id") <> "") Then 
  Recordset1__MMColParam = Request.QueryString("id")
End If
%>
<%
Dim Recordset1
Dim Recordset1_numRows

Set Recordset1 = Server.CreateObject("ADODB.Recordset")
Recordset1.ActiveConnection = MM_connnection_STRING
Recordset1.Source = "SELECT * FROM yuding WHERE id = " + Replace(Recordset1__MMColParam, "'", "''") + ""
Recordset1.CursorType = 0
Recordset1.CursorLocation = 2
Recordset1.LockType = 1
Recordset1.Open()

Recordset1_numRows = 0
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
                  <p><img src="image/line.gif" width="590" height="11"></p>
                  <form ACTION="<%=MM_editAction%>" METHOD="POST" name="form1">
                    <table width="630" align="center" cellpadding="0" cellspacing="0"bordercolor="#99BB99" style=" border-collapse: collapse">
                      <tr>
                        <td width=186 height="29" align="right" class="STYLE1"><p align="center"  class="greenb">商品编码                                                  
                        </td>
                        <td width=442 align="right"><div align="left">
                            <input name="sp_id" type="text" value="<%=(Recordset1.Fields.Item("sp_id").Value)%>" >
                        </div></td>
                      </tr>
                      <tr>
                        <td height="29" align="right" class="STYLE1"><div align="center">商品名称</div></td>
                        <td align="right"><div align="left">
                            <input name="sp_title" type="text" value="<%=(Recordset1.Fields.Item("sp_title").Value)%>" >
                        </div></td>
                      </tr>
                      <tr>
                        <td height="27" align="right" class="STYLE1"><div align="center"><span class="greenb">用户名</span></div></td>
                        <td align="right"><div align="left">
                            <input name="yd_personid" type="text" value="<%=(Recordset1.Fields.Item("yd_personid").Value)%>">
                        </div></td>
                      </tr>
                      <tr>
                        <td height="27" align="right" class="STYLE1"><div align="center">真实姓名</div></td>
                        <td align="right"><div align="left">
                            <input name="yd_name" type="text" value="<%=(Recordset1.Fields.Item("yd_name").Value)%>">
                        </div></td>
                      </tr>
                      <tr>
                        <td height="33" align="right" class="STYLE1"><div align="center"><span class="greenb">电话</span> </div></td>
                        <td align="right"><div align="left">
                            <input name="phone" type="text" value="<%=(Recordset1.Fields.Item("yd_phone").Value)%>">
                        </div></td>
                      </tr>
                      <tr>
                        <td height="23" align="right" class="STYLE1"><div align="center"><span class="greenb">地址</span></div></td>
                        <td align="right"><div align="left">
                            <input name="address" type="text" value="<%=(Recordset1.Fields.Item("yd_address").Value)%>">
                        </div></td>
                      </tr>
                      <tr>
                        <td height="24" align="right" class="STYLE1"><div align="center"><span class="greenb">汇款方式</span><span class="greenb"></span></div></td>
                        <td align="right"><div align="left">
                            <select name="fangshi" size="1">
                              <option value="汇款">汇款</option>
                              <option value="信誉卡">信誉卡</option>
                            </select>
                            </label>
                        </div></td>
                      </tr>
					        <tr>
                        <td width="186" height="28" align="center"><div align="center"><span class="STYLE1"><span class="greenb">是否成功</span></span></div></td>
                        <td width="442" align="center"><div align="left">
                          <label>
<input <%If (CStr((Recordset1.Fields.Item("yd_ifsuccess").Value)) = CStr("True")) Then Response.Write("checked=""checked""") : Response.Write("")%> name="checkbox" type="checkbox" value="1">                          
是</label>
                          <label></label>
                        </div></td>
                      </tr>
                      <tr>
                        <td width="186" height="32" align="center">
                            <input type="submit" name="Submit" value="提交">
                            </label>
                        </a></span></div></td>
                        <td width="442" align="center"><div align="left">
                            <label>
                            <input type="reset" name="Submit2" value="重置">
                            </label>
                        </div></td>
                      </tr>
                    </table>
                    <input type="hidden" name="MM_insert" value="form1">
                    <input type="hidden" name="MM_update" value="form1">
                    <input type="hidden" name="MM_recordId" value="<%= Recordset1.Fields.Item("id").Value %>">
                  </form>
                  <p><br>
                  </p>
                  <p><br>
                    <br>
                  </p></td>
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
<%
Recordset1.Close()
Set Recordset1 = Nothing
%>
