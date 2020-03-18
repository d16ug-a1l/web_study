<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>新建网页 2</title>
</head>

<body>
<%
ID=Trim(Request.QueryString("ID"))				'获取用户购买的商品序号
'判断商品序号是否为空，为空则停止处理
If ID="" Then 
	Response.write "没有选择商品！"
	Response.end
End If
ID=Cint(ID)								'把商品序号转换成数值
'获取用户购买数量
Num=Trim(Request.Form("Text"&ID))
'判断用户购买数量是否正确，不正确则停止处理
If Num="" Then
	Response.write "没有选择商品！"
	Response.end
End If
Num=Cint(Num)
Dim Total									'保存总价
'计算总价
Total=0
If ID=1 Then
	Total=Total+Num*1800
ElseIf ID=2 Then
	Total=Total+Num*120
ElseIf ID=3 Then
	Total=Total+Num*170
ElseIf ID=4 Then
	Total=Total+Num*130
End If
'判断总价是否为0。为0则停止处理
If Total=0 Then
	Response.write "没有选择商品！"
	Response.end
End If
 
Count=Session("Count")						'获取购买的商品次数
Session("Count")=Session("Count")+1
GWC=Session("GWCH")						'获取购买的商品序号
GWCTotal=Session("GWCHTotal")				'获取购买的商品的价格
GWC(Count+1)=ID							'添加新的商品序号
GWCTotal(Count+1)=Total					'添加新购买的商品价格

Session("GWCH")=GWC						'保存新购买的商品
Session("GWCHTotal")=GWCTotal				'保存新购买商品价格
GWCTotal=Session("GWCHTotal")	

%>
<p align="center"><a href="GWC.asp">查询购物车</a></p>
</body>

</html>
