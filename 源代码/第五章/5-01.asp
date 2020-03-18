<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>新建网页 1</title>
</head>
<body>
<%
'定义数组ASPContent并赋值
Dim ASPContent(4)
ASPContent(0)="ASP大全目录"
ASPContent(1)="ASP内置对象"
ASPContent(2)="ASPFSO对象"
ASPContent(3)="ASPADO对象"
ASPContent(4)="ASP网站安全"
'把ASPContent数组的值存在Session对象中
Session("ASP")=ASPContent
'从Session队象中获取存有的数组
strASP=Session("ASP")
'遍历数组中的所有元素并输出
for each str in strASP
	Response.write str&"<BR>"
Next
%>
</body>
</html>
