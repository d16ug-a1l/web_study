<!--#include file="top.asp"-->


<%session("new")="new"%>
<table width="98%" border="0" cellspacing="0" cellpadding="0" align="center"style="BORDER-LEFT: #000000 1px solid; BORDER-RIGHT: #000000 1px solid" bgcolor="#64B3D9" >
  <tr>
    <td> <SCRIPT language=JavaScript>
function FrontPage_Form1_Validator(theForm)
{
  if (theForm.name.value == "")
  {
    alert("�ܸ���������˭��");
    theForm.name.focus();
    return (false);
  }
  if (theForm.title.value == "")
  {
    alert("����Ҫ������˼��");
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
    alert("���������ַ����ȷ!");
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
       alert("���������ַ����ȷ!");
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
          <td colspan="4" height="25"><b><font color="#FFFFFF">�������������� </font></b> 
          </td>
        </tr>
        <tr> 
          <td width="16%" align="center" bgcolor="#92C8E2">��&nbsp; ��&nbsp; �ˣ�</td>
          <td colspan="2" bgcolor="#92C8E2"> 
            <input type="text" name="name" size="20" maxlength="15" >
              ��ID���� <b>5</b> ���������ڣ�<font color="#FF0000">*</font> </td>
        </tr>
        <tr> 
          <td width="16%" align="center" bgcolor="#DAEDF5">��&nbsp; �룺</td>
          <td colspan="2" bgcolor="#DAEDF5"> 
            <input type="password" name="pass" size="20" >
            ����д�����Ϊ�����û����������޷�ʹ������ID�� </td>
        </tr>
        <tr> 
          <td width="16%" align="center" bgcolor="#92C8E2">�� �� �� �⣺</td>
          <td colspan="2" bgcolor="#92C8E2"> 
            <input type="text" name="title" size="50" maxlength="40" >
            ���������� <b>20</b> ���������ڣ�<font color="#FF0000">*</font></td>
        </tr>
        <tr> 
          <td width="16%" align="center" bgcolor="#DAEDF5">�� �� �� ����</td>
          <td width="24%" bgcolor="#FFFFFF"> 
            <input type="text" name="email" size="20" >
          </td>
          <td width="60%" bgcolor="#DAEDF5">�� ҳ �� ַ�� 
            <input type="text" name="url" size="30"  value="http://">
          </td>
        </tr>
        <tr> 
          <td width="16%" align="center" bgcolor="#92C8E2">ѡ �� �� �飺</td>
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
              <option selected value="images/01.gif">�û�ͷ��-01 
              <option selected value="images/02.gif">�û�ͷ��-02 
              <option selected value="images/03.gif">�û�ͷ��-03 
              <option selected value="images/04.gif">�û�ͷ��-04 
              <option selected value="images/05.gif">�û�ͷ��-05 
              <option selected value="images/06.gif">�û�ͷ��-06 
              <option selected value="images/07.gif">�û�ͷ��-07 
              <option selected value="images/08.gif">�û�ͷ��-08 
              <option selected value="images/09.gif">�û�ͷ��-09 
              <option selected value="images/10.gif">�û�ͷ��-10 
              <option selected value="images/11.gif">�û�ͷ��-11 
              <option selected value="images/12.gif">�û�ͷ��-12 
              <option selected value="images/13.gif">�û�ͷ��-13 
              <option selected value="images/14.gif">�û�ͷ��-14 
              <option selected value="images/15.gif">�û�ͷ��-15 
              <option selected value="images/16.gif">�û�ͷ��-16 
              <option selected value="images/17.gif">�û�ͷ��-17 
              <option selected value="images/18.gif">�û�ͷ��-18 
              <option selected value="images/19.gif">�û�ͷ��-19 
              <option selected value="images/20.gif">�û�ͷ��-20 
            </select>
            <img id=idface src="images/01.gif" alt=�����������> </td>
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
              <input type="submit" value=" �ύ " name="B1" >
              &nbsp;&nbsp;&nbsp; 
              <input type="reset" value=" ��� " name="B2" >
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
