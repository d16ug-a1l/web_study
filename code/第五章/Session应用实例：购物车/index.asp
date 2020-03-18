<html>

<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>商品名称</title>
</head>

<body>
<%
Session("Count")=0			'设置购买商品的次数为0
Dim GWC()				'声明数组
Redim GWC(10)
'对数组的每个元素赋值
For i=0 to 10
GWC(i)=0
Next
Session("GWCH")=GWC			'把数组保存在Session变量中
Session("GWCHTotal")=GWC
%>

<div align="center">

<table   width="53%" id="table1">
	<tr>
		<td width="181">商品名称</td>
		<td width="103">价格</td>
		<td width="67">数量</td>
		<td>操作</td>
	</tr>
	<tr>
		<td width="181">液晶显示器</td>
		<td>1800</td>
		<form method="post" action="Insert.asp?id=1">
		<td width="67"><input type=text name="Text1" size="8"></td>
		<td width="79"><input type=submit name="提交"></td>
		</form>
	</tr>	
	<tr>
		<td width="181">键盘</td>
		<td>120</td>
		<form method="post" action="Insert.asp?id=2">
		<td width="67"><input type=text name="Text2" size="8"></td>
		<td width="79"><input type=submit name="提交"></td>
		</form>
	</tr>	
	<tr>
		<td width="181">1G优盘</td>
		<td>170</td>
		<form method="post" action="Insert.asp?id=3">
		<td width="67"><input type=text name="Text3" size="8"></td>
		<td width="79"><input type=submit name="提交"></td>
		</form>
	</tr>
	<tr>
		<td width="181">光电鼠标</td>
		<td>130</td>
		<form method="post" action="Insert.asp?id=4">
		<td width="67"><input type=text name="Text4" size="8"></td>
		<td width="79"><input type=submit name="提交"></td>
		</form>
	</tr>
</table>

</div>

<p align="center"><a href="GWC.asp">查询购物车</a></p>

</body>

</html>