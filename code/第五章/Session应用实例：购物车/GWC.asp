<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�½���ҳ 2</title>
</head>

<body>
<%
Num=Cint(Num)
Dim Total
Total=0
'�����ȡ�������Ʒ��źͼ۸�
Count=Session("Count")
GWC=Session("GWCH")
GWCHTotal=Session("GWCHTotal")
'ѭ����ʾ���еĹ��ﳵ��Ϣ
str=""
 
for i=1 To Count
	ID=GWC(i)
	Total=GWCHTotal(i)
    If ID=1 Then
		str= "Һ����ʾ��������Ϊ1800���ܼ�Ϊ"&Total
	ElseIf ID=2 Then
		str= "���̡�����Ϊ120���ܼ�Ϊ"&Total
	ElseIf ID=3 Then
		str= "1G���̡�����Ϊ170���ܼ�Ϊ"&Total
	ElseIf ID=4 Then
		str= "�����ꡣ����Ϊ130���ܼ�Ϊ"&Total
	End If
	Response.write str&"<BR>"
Next

%>
</body>

</html>
