<!--#include file="top.asp"-->
<%
     if request("action")="save" then
      
      title=Trim(Request.form("title"))
      user_name=Trim(Request.form("name"))
      pass=Trim(Request.form("pass"))
      nei=Trim(Request.form("nei"))
      
      

	bad=split(badstr,"|")
	for i=0 to UBound(bad)
		nei=Replace(nei,bad(i),"**")
	next
	title=server.htmlencode(title)
		bad=split(badstr,"|")
	for i=0 to UBound(bad)
		title=Replace(title,bad(i),"**")
	next
	user_name=server.htmlencode(user_name)
		bad=split(badstr,"|")
	for i=0 to UBound(bad)
		user_name=Replace(user_name,bad(i),"**")
	next


if user_name="" or pass="" then
response.write"<SCRIPT language=JavaScript>alert('请填写您的帐号、密码');"
response.write"javascript:history.go(-1)</SCRIPT>"
response.end
end if

sql="select * from user where name='"&user_name&"' and pass='"&pass&"'"
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.open sql,conn,3,3
if rs.eof or rs.bof then
response.write"<SCRIPT language=JavaScript>alert('您不能发布公告');"
response.write"javascript:history.go(-1)</SCRIPT>"
response.end
end if
session("user_name")=rs("name")
session("bz")=rs("bz")
rs.close
set rs=nothing

sql="select * from smallpager "
Set rs = Server.CreateObject("ADODB.Recordset")
 rs.Open sql,conn,3,3
     rs.addnew
     rs("user_name")=user_name
     rs("s_title")=title
     rs("adate")=now()
	 rs("hit")=1
     rs.update
	 rs.close
	 set rs=nothing
response.redirect "index.asp" 
end if
   
%>
<table width="98%" border="0" cellspacing="0" cellpadding="0" align="center"style="BORDER-LEFT: #000000 1px solid; BORDER-RIGHT: #000000 1px solid" bgcolor="#64B3D9" >
  <tr>
    <td> <SCRIPT language=JavaScript>
function FrontPage_Form1_Validator(theForm)
{
  if (theForm.name.value == "")
  {
    alert("请问您是谁？");
    theForm.name.focus();
    return (false);
  }
  if (theForm.pass.value == "")
  {
    alert("请写上您的密码!");
    theForm.pass.focus();
    return (false);
  }
if (theForm.title.value == "")
  {
    alert("发言要有中心思想");
    theForm.title.focus();
    return (false);
  }
 

  return (true);
}
</script>
      <form method="POST" action="smallpaper.asp?action=save" name="form1"  onsubmit="return FrontPage_Form1_Validator(this)">
        <br>
        <br>
        <table border="0" width="98%" cellspacing="1" cellpadding="3" bordercolorlight="#FFE8E8" bordercolordark="#FFFFFF" align="center" bgcolor="#205E7B">
          <tr> 
            <td colspan="4" height="25"><b><font color="#FFFFFF">　　发 布 公 告</font></b> 
            </td>
        </tr>
        <tr> 
            <td width="23%" align="center" bgcolor="#92C8E2">发&nbsp;布 人：</td>
            <td colspan="2" bgcolor="#92C8E2" width="77%"> 
              <input type="text" name="name" size="20" maxlength="15" >
            （只有保留用户才能发布小字报）<font color="#FF0000">*</font> </td>
        </tr>
        <tr> 
            <td width="23%" align="center" bgcolor="#DAEDF5">密&nbsp; 码：</td>
            <td colspan="2" bgcolor="#DAEDF5" width="77%"> 
              <input type="password" name="pass" size="20" ><font color="#FF0000">*</font>
             </td>
        </tr>
        <tr> 
            <td width="23%" align="center" bgcolor="#92C8E2">内容：</td>
            <td colspan="2" bgcolor="#92C8E2" width="77%"> 
              <input type="text" name="title" size="50" maxlength="40" >
            （内容限制 <b>20</b> 个汉字以内）<font color="#FF0000">*</font></td>
        </tr>
        
        
        <tr> 
          <td colspan="3" > 
            <p align="center"> 
              <input type="submit" value=" 提交 " name="B1" >
              &nbsp;&nbsp;&nbsp; 
              <input type="reset" value=" 清除 " name="B2" >
          </td>
        </tr>
      </table>
    </form></td>
  </tr>
</table>
<table width="98%" border="0" cellspacing="0" cellpadding="0" align="center" bgcolor="#000000">
  <tr> 
    <td align="center" height="30" class=td1> 
      <%=footer%>
    </td>
  </tr>
</table>
</body>
</html>
