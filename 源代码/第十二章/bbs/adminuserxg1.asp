<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>�ޱ����ĵ�</title>
<style type="text/css">
<!--
.zt {
	font-family: "����";
	font-size: 14px;
	color: #205E7B;
}
.STYLE1 {color: #FFFFFF}
-->
</style>
</head>

<body>
<table width="95%" border="0" cellspacing="1" cellpadding="0" align="center">
  <tr>
    <td height="50" class="zt">�˳�&nbsp;&nbsp;����Ա�ʺŹ���&nbsp;&nbsp;���ӹ���&nbsp;&nbsp;</td>
  </tr>
</table>
<br />
<table width="96%" border="0" align="center" cellspacing="1" bgcolor="#205E7B">
  <tr align="center">
    <td  height="25"><div align="left" class="STYLE1">�޸��ʺź�����</div></td>
  </tr>
  <tr>
    <td height="1" colspan="3" bgcolor="#000000"></td>
  </tr>
<tr>
          <td height="35" colspan="7" bgcolor="#D6DFF7"><form  name="form1">
            <table width="96%" border="0" cellspacing="1" cellpadding="3" align="center">
              <tr>
                <td>�ʺ����ƣ�
                  <label>
                  <input name="username" type="text" />
                </label></td>
              </tr>
              <tr>
                <td>�޸����룺
                  <input type="text" name="pass" value="" /></td>
              </tr>
              <tr>
                <td>�ظ����룺
                  <input type="text" name="pass1" value="" /></td>
              </tr>
              <tr>
                <td><input type="submit" name="admin" value="ȷ��" onclick="check()" />
                &nbsp;
                <label>
                <input type="reset" name="button" id="button" value="����" />
                &nbsp;&nbsp;����</label></td>
              </tr>
            </table>
          
          
            
            
</form></td>
  </tr>
</table>
<script language="javascript">
function check(){
if(document.form1.pass.value==""||document.form1.pass1.value=="")
{alert("����������")}
else{
if(document.form1.pass.value!=document.form1.pass1.value)
{
alert("�����������벻��ȷ")
}
}
}

</script>
</body>
</html>