<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>新建网页 1</title>
</head>

<body>
<% 
'求出该文件的物理路径，以确定保存图像的文件夹的物理路径
path=server.mappath("savephoto.asp")
'path路径去掉文件名就是保存图像的文件夹物理路径
'反向查找"\"的位置n，截取path的前n位就是保存图像的文件夹路径
path=mid(path,1,instrRev(path,"\"))
'获取所有请求的信息并存放入Stream对象中
formsize=request.totalbytes '求出整个接受的数据大小
formdata=request.binaryread(formsize)'读取整个二进制数据
'创建Stream对象并打开
Set st1=Server.CreateObject("Adodb.Stream")
st1.Type= 1
st1.Mode=3
st1.open
'把获取的请求数据写入st1对象中，以供以后截取使用
st1.Write formdata
return=chrB(13)&chrB(10)  '构造一个回车换行符号
'求出起始标志数据。起始标志数据就是第一个回车换行符之前的数据
divider=leftB(formdata,clng(instrb(formdata,return))-1) 
'下面是求出第二个起始标志起始位置
datastart=instrb(lenb(divider),formdata,divider)
'下面是求出第三个起始标志起始位置 
datastart1=instrb(datastart+1,formdata,divider) 
'下面是求出文件数据的起始位置。
'文件数据起始位置等于第三个起始标志后的回车换行符的起始位置加上回车换行符的字节数
datastart=instrb(datastart1+1,formdata,return&return) +3
'获取文件路径。把第三个起始标志后的文件路径信息转换成文本，通过查找“finlename=”获取文件路径信息
set tempStream = Server.CreateObject("adodb.stream")
tempStream.Type = 1
tempStream.Mode =3
tempStream.Open
'把st1中文件路径信息部分数据复制到tempStream对象中
st1.Position=datastart1
st1.CopyTo tempStream ,datastart-datastart1 
'把tempStream对象中的数据以文本形式读取
tempStream.Position = 0
tempStream.Type = 2
tempStream.Charset ="gb2312"
'FilePath中保存文件路径信息的文本形式
FilePath= tempStream.ReadText
'确定”filename=”的位置，通过查找“.”获取文件扩展名起始位置
Filename=mid(FilePath,instr(FilePath,"filename=")+10)
'获取扩展名，并组合出文件的路径和文件名
ext=mid(Filename,instr(Filename,"."),4)
'makefilename是依据当前时间组合出文件名的函数
filename="images\"&makefilename&ext
filename=path&filename
'求出图片文件的长度，文件的结束位置就是第四个起始标志的开始处
'因此文件的长度就是第四个起始标志的开始位置减去文件信息的起始位置
dataend=instrb(datastart+1,formdata,divider)-datastart 
Set st2=Server.CreateObject("Adodb.Stream")
st2.Type= 1
st2.Mode=3
st2.open
st1.Position=datastart
'把st1对象中的从datastart开始的dataend个字节数据复制到st2对象中
st1.copyto st2,dataend
'把st2对象的数据保存在文件filename中
st2.SaveToFile filename,2
st2.Close
response.write "<H2>图片上传成功！</H2>"

%>

</body>

</html>
