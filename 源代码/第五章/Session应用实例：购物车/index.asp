<html>

<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>��Ʒ����</title>
</head>

<body>
<%
Session("Count")=0			'���ù�����Ʒ�Ĵ���Ϊ0
Dim GWC()				'��������
Redim GWC(10)
'�������ÿ��Ԫ�ظ�ֵ
For i=0 to 10
GWC(i)=0
Next
Session("GWCH")=GWC			'�����鱣����Session������
Session("GWCHTotal")=GWC
%>

<div align="center">

<table   width="53%" id="table1">
	<tr>
		<td width="181">��Ʒ����</td>
		<td width="103">�۸�</td>
		<td width="67">����</td>
		<td>����</td>
	</tr>
	<tr>
		<td width="181">Һ����ʾ��</td>
		<td>1800</td>
		<form method="post" action="Insert.asp?id=1">
		<td width="67"><input type=text name="Text1" size="8"></td>
		<td width="79"><input type=submit name="�ύ"></td>
		</form>
	</tr>	
	<tr>
		<td width="181">����</td>
		<td>120</td>
		<form method="post" action="Insert.asp?id=2">
		<td width="67"><input type=text name="Text2" size="8"></td>
		<td width="79"><input type=submit name="�ύ"></td>
		</form>
	</tr>	
	<tr>
		<td width="181">1G����</td>
		<td>170</td>
		<form method="post" action="Insert.asp?id=3">
		<td width="67"><input type=text name="Text3" size="8"></td>
		<td width="79"><input type=submit name="�ύ"></td>
		</form>
	</tr>
	<tr>
		<td width="181">������</td>
		<td>130</td>
		<form method="post" action="Insert.asp?id=4">
		<td width="67"><input type=text name="Text4" size="8"></td>
		<td width="79"><input type=submit name="�ύ"></td>
		</form>
	</tr>
</table>

</div>

<p align="center"><a href="GWC.asp">��ѯ���ﳵ</a></p>

</body>

</html>