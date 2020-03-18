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
' *** Insert Record: set variables

If (CStr(Request("MM_insert")) = "form1") Then

  MM_editConnection = MM_connnection_STRING
  MM_editTable = "yuding"
  MM_editRedirectUrl = "link.html"
  MM_fieldsStr  = "sp_id|value|sp_title|value|yd_personid|value|yd_name|value|phone|value|address|value|fangshi|value|time|value"
  MM_columnsStr = "sp_id|',none,''|sp_title|',none,''|yd_personid|',none,''|yd_name|',none,''|yd_phone|',none,''|yd_address|',none,''|fanshi|',none,''|yd_time|',none,NULL"

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
' *** Insert Record: construct a sql insert statement and execute it

Dim MM_tableValues
Dim MM_dbValues

If (CStr(Request("MM_insert")) <> "") Then

  ' create the sql insert statement
  MM_tableValues = ""
  MM_dbValues = ""
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
      MM_tableValues = MM_tableValues & ","
      MM_dbValues = MM_dbValues & ","
    End If
    MM_tableValues = MM_tableValues & MM_columns(MM_i)
    MM_dbValues = MM_dbValues & MM_formVal
  Next
  MM_editQuery = "insert into " & MM_editTable & " (" & MM_tableValues & ") values (" & MM_dbValues & ")"

  If (Not MM_abortEdit) Then
    ' execute the insert
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
<html>
<head>
<title>Freedom</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type="text/css">
<!--
.STYLE1 {	color: #0000FF;
	font-size: 14px;
}
-->
</style>
</head>

<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" text="#333333" onLoad="MM_preloadImages('image/bt_01_on.gif','image/bt_02_on.gif','image/bt_04_on.gif','image/bt_05_on.gif')" link="#6633FF" vlink="#6666FF" alink="#6666FF">
<table width="100%" border="0" cellspacing="0" cellpadding="0" height="100%">
  <tr>
    <td height="487" valign="top"> 
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td background="image/bg_top2.gif" height="67" valign="top"> 
            <table width="536" border="0" cellspacing="0" cellpadding="0" background="">
              <tr> 
                <td width="190"><img src="image/transparent.gif" width="10" height="20"></td>
                <td width="346" align="right"><a href="../admin/login.asp"><img src="image/con_contact.gif" width="66" height="15" border="0"></a></td>
              </tr>
              <tr> 
                <td width="190"><img src="image/transparent.gif" width="10" height="22"></td>
                <td width="346">　</td>
              </tr>
              <tr>
                <td width="190">
                  <table border="0" cellspacing="0" cellpadding="0">
                    <tr>
                      <td><img src="image/obj_title.gif" width="190" height="28"></td>
                    </tr>
                    <tr>
                      <td><img src="image/transparent.gif" width="10" height="1"></td>
                    </tr>
                  </table>
                </td>
                <td width="346" valign="bottom"> 
                  <table border="0" cellspacing="0" cellpadding="0">
                    <tr> 
                      <td><a href="../index.html" ><img src="image/bt_01_off.gif" width="86" height="18" border="0" name="Image5"></a></td>
                      <td><a href="index.asp" ><img src="image/bt_02_off.gif" width="86" height="18" border="0" name="Image1"></a></td>
                      <td><img src="image/bt_03_on.gif" width="86" height="18" border="0" name="Image2"></td>
                      <td><a href="../html/chaxun1.asp" ><img src="image/bt_04_off.gif" width="86" height="18" border="0" name="Image3"></a></td>
                      <td><a href="link.html" ><img src="image/bt_05_off.gif" width="86" height="18" border="0" name="Image4"></a></td>
                      <td><img src="image/obj_bt.gif" width="9" height="18"></td>
                    </tr>
                  </table>
                </td>
              </tr>
              <tr> 
                <td width="190"><img src="image/transparent.gif" width="10" height="25"></td>
                <td width="346">　</td>
              </tr>
            </table>
          </td>
        </tr>
        <tr>
          <td height="5" valign="top"><img src="image/transparent.gif" width="20" height="10"></td>
        </tr>
      </table>
      <table width="640" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="16"><img src="image/transparent.gif" width="50" height="20"></td>
          <td width="618" valign="top"> 
            <table border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td> 
                  <table width="583" border="0" cellspacing="0" cellpadding="0">
                    <tr> 
                      <td width="583">&nbsp;</td>
                    </tr>
                    <tr> 
                      <td><form ACTION="<%=MM_editAction%>" METHOD="POST" name="form1">
                        <table width="630" align="center" cellpadding="0" cellspacing="0"bordercolor="#99BB99" style=" border-collapse: collapse">
                          <tr>
                            <td width=186 height="29" align="right" class="STYLE1"><p align="center"  class="greenb">商品编码                                                  
                            </td>
                            <td width=442 align="right"><div align="left">
                                <input type="text" name="sp_id" value="<%=request.QueryString("id")%>">
                            </div></td>
                          </tr>
                          <tr>
                            <td height="29" align="right" class="STYLE1"><div align="center">商品名称</div></td>
                            <td align="right"><div align="left">
                                <input type="text" name="sp_title" value="<%=request.QueryString("title")%>">
                            </div></td>
                          </tr>
                          <tr>
                            <td height="27" align="right" class="STYLE1"><div align="center"><span class="greenb">用户名</span></div></td>
                            <td align="right"><div align="left">
                                <input type="text" name="yd_personid">
                            </div></td>
                          </tr>
						                            <tr>
                            <td height="27" align="right" class="STYLE1"><div align="center">真实姓名</div></td>
                            <td align="right"><div align="left">
                                <input type="text" name="yd_name">
                            </div></td>
                          </tr>
                          <tr>
                            <td height="33" align="right" class="STYLE1"><div align="center"><span class="greenb">电话</span> </div></td>
                            <td align="right"><div align="left">
                                <input type="text" name="phone">
                            </div></td>
                          </tr>
                          <tr>
                            <td height="23" align="right" class="STYLE1"><div align="center"><span class="greenb">地址</span></div></td>
                            <td align="right"><div align="left">
                                <input type="text" name="address">
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
                            <td width="186" height="28" align="center"><div align="center"><span class="STYLE1"><span class="greenb">上传时间</span></span></div></td>
                            <td width="442" align="center"><div align="left">
                                <input name="time" type="text" value="<%=now()%>">
                            </div></td>
                          </tr>
                          <tr>
                            <td width="186" height="32" align="center"><div align="center"><span class="STYLE1"><a href="zhuce.asp">
                                <label>
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
                      </form>
                      </td>
                    </tr>
                  </table>
                  
                  
                  
                </td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
    <div align="center"></div></td>
  </tr>
  <tr>
    <td valign="bottom"> 
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td background="image/bg_bottom.gif" align="right">
            <table border="0" cellspacing="0" cellpadding="0" background="">
              <tr> 
                <td height="14"><img src="image/transparent.gif" width="210" height="20"></td>
                <td rowspan="2"><img src="image/transparent.gif" width="3" height="40"></td>
              </tr>
              <tr> 
                <td><font size="1" color="#FFFFFF"></td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
    </td>
  </tr>
</table>
</body>
</html>