<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�½���ҳ 1</title>
</head>

<body>
<% 
'������ļ�������·������ȷ������ͼ����ļ��е�����·��
path=server.mappath("savephoto.asp")
'path·��ȥ���ļ������Ǳ���ͼ����ļ�������·��
'�������"\"��λ��n����ȡpath��ǰnλ���Ǳ���ͼ����ļ���·��
path=mid(path,1,instrRev(path,"\"))
'��ȡ�����������Ϣ�������Stream������
formsize=request.totalbytes '����������ܵ����ݴ�С
formdata=request.binaryread(formsize)'��ȡ��������������
'����Stream���󲢴�
Set st1=Server.CreateObject("Adodb.Stream")
st1.Type= 1
st1.Mode=3
st1.open
'�ѻ�ȡ����������д��st1�����У��Թ��Ժ��ȡʹ��
st1.Write formdata
return=chrB(13)&chrB(10)  '����һ���س����з���
'�����ʼ��־���ݡ���ʼ��־���ݾ��ǵ�һ���س����з�֮ǰ������
divider=leftB(formdata,clng(instrb(formdata,return))-1) 
'����������ڶ�����ʼ��־��ʼλ��
datastart=instrb(lenb(divider),formdata,divider)
'�����������������ʼ��־��ʼλ�� 
datastart1=instrb(datastart+1,formdata,divider) 
'����������ļ����ݵ���ʼλ�á�
'�ļ�������ʼλ�õ��ڵ�������ʼ��־��Ļس����з�����ʼλ�ü��ϻس����з����ֽ���
datastart=instrb(datastart1+1,formdata,return&return) +3
'��ȡ�ļ�·�����ѵ�������ʼ��־����ļ�·����Ϣת�����ı���ͨ�����ҡ�finlename=����ȡ�ļ�·����Ϣ
set tempStream = Server.CreateObject("adodb.stream")
tempStream.Type = 1
tempStream.Mode =3
tempStream.Open
'��st1���ļ�·����Ϣ�������ݸ��Ƶ�tempStream������
st1.Position=datastart1
st1.CopyTo tempStream ,datastart-datastart1 
'��tempStream�����е��������ı���ʽ��ȡ
tempStream.Position = 0
tempStream.Type = 2
tempStream.Charset ="gb2312"
'FilePath�б����ļ�·����Ϣ���ı���ʽ
FilePath= tempStream.ReadText
'ȷ����filename=����λ�ã�ͨ�����ҡ�.����ȡ�ļ���չ����ʼλ��
Filename=mid(FilePath,instr(FilePath,"filename=")+10)
'��ȡ��չ��������ϳ��ļ���·�����ļ���
ext=mid(Filename,instr(Filename,"."),4)
'makefilename�����ݵ�ǰʱ����ϳ��ļ����ĺ���
filename="images\"&makefilename&ext
filename=path&filename
'���ͼƬ�ļ��ĳ��ȣ��ļ��Ľ���λ�þ��ǵ��ĸ���ʼ��־�Ŀ�ʼ��
'����ļ��ĳ��Ⱦ��ǵ��ĸ���ʼ��־�Ŀ�ʼλ�ü�ȥ�ļ���Ϣ����ʼλ��
dataend=instrb(datastart+1,formdata,divider)-datastart 
Set st2=Server.CreateObject("Adodb.Stream")
st2.Type= 1
st2.Mode=3
st2.open
st1.Position=datastart
'��st1�����еĴ�datastart��ʼ��dataend���ֽ����ݸ��Ƶ�st2������
st1.copyto st2,dataend
'��st2��������ݱ������ļ�filename��
st2.SaveToFile filename,2
st2.Close
response.write "<H2>ͼƬ�ϴ��ɹ���</H2>"

%>

</body>

</html>
