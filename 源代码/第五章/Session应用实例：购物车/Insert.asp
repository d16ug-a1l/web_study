<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�½���ҳ 2</title>
</head>

<body>
<%
ID=Trim(Request.QueryString("ID"))				'��ȡ�û��������Ʒ���
'�ж���Ʒ����Ƿ�Ϊ�գ�Ϊ����ֹͣ����
If ID="" Then 
	Response.write "û��ѡ����Ʒ��"
	Response.end
End If
ID=Cint(ID)								'����Ʒ���ת������ֵ
'��ȡ�û���������
Num=Trim(Request.Form("Text"&ID))
'�ж��û����������Ƿ���ȷ������ȷ��ֹͣ����
If Num="" Then
	Response.write "û��ѡ����Ʒ��"
	Response.end
End If
Num=Cint(Num)
Dim Total									'�����ܼ�
'�����ܼ�
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
'�ж��ܼ��Ƿ�Ϊ0��Ϊ0��ֹͣ����
If Total=0 Then
	Response.write "û��ѡ����Ʒ��"
	Response.end
End If
 
Count=Session("Count")						'��ȡ�������Ʒ����
Session("Count")=Session("Count")+1
GWC=Session("GWCH")						'��ȡ�������Ʒ���
GWCTotal=Session("GWCHTotal")				'��ȡ�������Ʒ�ļ۸�
GWC(Count+1)=ID							'����µ���Ʒ���
GWCTotal(Count+1)=Total					'����¹������Ʒ�۸�

Session("GWCH")=GWC						'�����¹������Ʒ
Session("GWCHTotal")=GWCTotal				'�����¹�����Ʒ�۸�
GWCTotal=Session("GWCHTotal")	

%>
<p align="center"><a href="GWC.asp">��ѯ���ﳵ</a></p>
</body>

</html>
