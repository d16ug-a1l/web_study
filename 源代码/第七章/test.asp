<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>新建网页 1</title>
</head>

<body>
<%
formsize=request.totalbytes '求出整个数据的大小
formdata=request.binaryread(formsize)'读取整个二进制数据
return=chrB(13)&chrB(10)
divider=leftB(formdata,clng(instrb(formdata,return))-1) 
datastart=instrb(lenb(divider),formdata,divider)

datastart1=instrb(datastart+1,formdata,divider)   
datastart=instrb(datastart1+1,formdata,return&return) +3
set tempStream = Server.CreateObject("adodb.stream")
tempStream.Type = 1
tempStream.Mode =3
tempStream.Open
st1.Position=datastart1
st1.CopyTo tempStream ,datastart-datastart1 
tempStream.Position = 0
tempStream.Type = 2
tempStream.Charset ="gb2312"
FilePath= tempStream.ReadText

Filename=mid(FilePath,instr(FilePath,"filename=")+10)
'ext=mid(Filename,instr(Filename,"."),4)
'filename1=makefilename&ext
dataend=instrb(datastart+1,formdata,divider)-datastart '求出图片的长度

mydata=midb(formdata,datastart,dataend) 

Set Conn= Server.CreateObject("ADODB.Connection")
Conn.ConnectionString="Provider=Microsoft.Jet.OLEDB.4.0;"&_
			 			"Data Source="&Server.MapPath("img.mdb")
Conn.Open
Set rs=Server.CreateObject("ADODB.Recordset")
TableName="[image]"
rs.Open TableName,Conn,1,3
rs.AddNew
rs("IMG").AppendChunk mydata
rs("FilePath")=Filename
rs("FileName")=Name
rs.Update
set rs=nothing
set Conn=Nothing
%>

</body>

</html>
