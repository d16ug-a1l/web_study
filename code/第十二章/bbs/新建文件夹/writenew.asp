<!--#include file="top.asp"-->


<%session("new")="new"%>
<table width="98%" border="0" cellspacing="0" cellpadding="0" align="center"style="BORDER-LEFT: #000000 1px solid; BORDER-RIGHT: #000000 1px solid" bgcolor="#64B3D9" >
  <tr>
    <td> <SCRIPT language=JavaScript>
function FrontPage_Form1_Validator(theForm)
{
  if (theForm.name.value == "")
  {
    alert("能告诉我你是谁吗？");
    theForm.name.focus();
    return (false);
  }
  if (theForm.title.value == "")
  {
    alert("发言要有中心思想");
    theForm.title.focus();
    return (false);
  }

  var checkOK = "0123456789abcdefghijklmnopqrstuvwxyz@._";
  var checkStr = theForm.email.value;
  var allValid = true;
  var decPoints = 0;
  var allNum = "";
  for (i = 0;  i < checkStr.length;  i++)
  {
    ch = checkStr.charAt(i);
    for (j = 0;  j < checkOK.length;  j++)
      if (ch == checkOK.charAt(j))
        break;
    if (j == checkOK.length)
    {
      allValid = false;
      break;
    }
    if (ch != ".")
      allNum += ch;
  }
  if (!allValid)
  {
    alert("您的信箱地址不正确!");
    theForm.email.focus();
    return (false);
  }
  

  if (theForm.email.value !== "")
  {
     var checkOK2 = theForm.email.value;
     var checkStr2 = "@.";
     var allValid2 = true;
     var decPoints2 = 0;
     var allNum2 = "";
     for (i = 0;  i < checkStr2.length;  i++)
     {
       ch2 = checkStr2.charAt(i);
       for (j = 0;  j < checkOK2.length;  j++)
         if (ch2 == checkOK2.charAt(j))
           break;
       if (j == checkOK2.length)
       {
         allValid2 = false;
         break;
       }
       if (ch2 != ".")
         allNum2 += ch2;
     }
     if (!allValid2)
     {
       alert("您的信箱地址不正确!");
       theForm.email.focus();
       return (false);
     }
  }
  

  return (true);
}
</script>
      <form method="POST" action="addnews.asp" name="form1"  onsubmit="return FrontPage_Form1_Validator(this)">
        <br>
        <table border="0" width="98%" cellspacing="1" cellpadding="3" bordercolorlight="#FFE8E8" bordercolordark="#FFFFFF" align="center" bgcolor="#205E7B">
        <tr> 
          <td colspan="4" height="25"><b><font color="#FFFFFF">　　发表新帖子 </font></b> 
          </td>
        </tr>
        <tr> 
          <td width="16%" align="center" bgcolor="#92C8E2">发&nbsp; 表&nbsp; 人：</td>
          <td colspan="2" bgcolor="#92C8E2"> 
            <input type="text" name="name" size="20" maxlength="15" >
              （ID限制 <b>5</b> 个汉字以内）<font color="#FF0000">*</font> </td>
        </tr>
        <tr> 
          <td width="16%" align="center" bgcolor="#DAEDF5">密&nbsp; 码：</td>
          <td colspan="2" bgcolor="#DAEDF5"> 
            <input type="password" name="pass" size="20" >
            （填写密码成为保留用户，其他人无法使用您的ID） </td>
        </tr>
        <tr> 
          <td width="16%" align="center" bgcolor="#92C8E2">帖 子 主 题：</td>
          <td colspan="2" bgcolor="#92C8E2"> 
            <input type="text" name="title" size="50" maxlength="40" >
            （标题限制 <b>20</b> 个汉字以内）<font color="#FF0000">*</font></td>
        </tr>
        <tr> 
          <td width="16%" align="center" bgcolor="#DAEDF5">电 子 邮 件：</td>
          <td width="24%" bgcolor="#FFFFFF"> 
            <input type="text" name="email" size="20" >
          </td>
          <td width="60%" bgcolor="#DAEDF5">主 页 地 址： 
            <input type="text" name="url" size="30"  value="http://">
          </td>
        </tr>
        <tr> 
          <td width="16%" align="center" bgcolor="#92C8E2">选 择 表 情：</td>
          <td colspan="2" bgcolor="#92C8E2"> 
            <% for i=0 to 17 %>
            <input type="radio" value="<%=(i)%>" name="pic" <%if i=0 then response.write "checked"%>>
            <img src="<%="title/face" &(i) & ".gif"%>" >&nbsp;&nbsp; 
            <%if i>0 and ((i+1) mod 9=0) then response.write "<br>"%>
            <%next%>
          </td>
        </tr>
        <tr> 
          <td width="16%" align="center" bgcolor="#DAEDF5"> 
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
            </select>
            <img id=idface src="images/01.gif" alt=个人形象代表> </td>
            <td valign="top" colspan="2" bgcolor="#DAEDF5">
              <table width="98%" border="0" cellspacing="0" cellpadding="0">
                <tr> 
                  <td> 
                    
<textarea name="nei" style="display:none"></textarea>
<iframe ID="bbs" src="edit/ewebeditor.asp?id=nei&style=bbs" frameborder="0" scrolling="no" width="98%" HEIGHT="350"></iframe> 



                  </td>
                  <td> </td>
                </tr>
              </table>
                </td>
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
    <td align="center" height="30"class=td1> 
      <%=footer%>
    </td>
  </tr>
</table>

 
        </body>
</html>
