<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>新建网页 2</title>
</head>

<body>
<%
Num=Cint(Num)
Dim Total
Total=0
'下面获取购买的商品序号和价格
Count=Session("Count")
GWC=Session("GWCH")
GWCHTotal=Session("GWCHTotal")
'循环显示所有的购物车信息
str=""
 
for i=1 To Count
	ID=GWC(i)
	Total=GWCHTotal(i)
    If ID=1 Then
		str= "液晶显示器。单价为1800。总价为"&Total
	ElseIf ID=2 Then
		str= "键盘。单价为120。总价为"&Total
	ElseIf ID=3 Then
		str= "1G优盘。单价为170。总价为"&Total
	ElseIf ID=4 Then
		str= "光电鼠标。单价为130。总价为"&Total
	End If
	Response.write str&"<BR>"
Next

%>
</body>

</html>
