<html>
<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>���������ʾ���������</title>
</head>
<body>
<%
Dim i
Dim Obj
'����Content Linking���
Set Obj=Server.CreateObject("MSWC.NextLink")
'��ȡ��ǰҳ����NextLink.txt�ļ��е����
i=Obj.GetListIndex("NextLink.txt")
Dim strContent
'��ȡ��ǰҳ�����������
strContent=Obj.GetNthDescription("NextLink.txt",i)
'�����ǰҳ�����������
Response.write("<p align='center'><font face='�����п�' size='6' color='#0000FF'>"& strContent&"</font></p> ")
Dim strPrev,strPrevURL
Dim strNext,strNextURL
If i=1 Then 
'�����ǰҳΪ��һҳ����û��ǰһҳ����
'û�д��жϣ��������ִ���
	strPrevURL=""
	strPrev="ǰһҳ"
Else
'��ȡǰһҳ��URL
strPrevURL=Obj.GetPreviousURL("NextLink.txt")
'��ȡǰһҳ����������
strPrev=Obj.GetPreviousDescription("NextLink.txt")
'����ǰһҳ������
strPrevURL="<a href='"& strPrevURL&"'>"
End If
If i=Obj.GetListCount("NextLink.txt") Then 
'�����ǰҳΪ���һҳ����û�к�һҳ����
'���û�д��жϣ�����ִ������
	strNext="��һҳ"
	strNextURL=""
Else
'��ȡ��һҳ����������
strNext=Obj.GetNextDescription("NextLink.txt")
'��ȡ��һҳ������URL
strNextURL=Obj.GetNextURL("NextLink.txt")
'���ú�һҳ������
strNextURL=" <a href='"& strNextURL&"'>"
End If
Response.write("<p align='center'>"&strPrevURL&strPrev&"</a>")
Response.write(strNextURL&strNext&"</a></p> ")
%>
<p>&nbsp;&nbsp;&nbsp;
<span lang="EN-US" style="font-size: 10.5pt; font-family: Times New Roman">AD 
Rotator</span><span style="font-size: 10.5pt; font-family: ����">�����һ�����������ʾ�����������AD Rotato����ܹ�ʵ�ֹ��ͼƬ�����ʾ��ҳ�ϣ����ҿ����趨���������ֵ�Ƶ�ʡ�AD Rotator���������������ӻ����޸Ĺ��ĳ����ӣ������û��Ϳ���ͨ�����������ʹ��ͻ���WEBվ�㡣AD Rotator�������������ʾÿ���������񣬲�������Ӻ��޸Ĺ��Ĺ�����ø����ɡ�</span></p>
</body>
</html>
