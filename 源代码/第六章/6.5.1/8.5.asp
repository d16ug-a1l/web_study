<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�½���ҳ 1</title>
</head>

<body>
<%
function GetBytes(Path, offset, bytes) 
	Dim objFSO 
	Dim objFTemp 
	Dim objTextStream 
	Dim lngSize 
	Set objFSO = CreateObject("ADODB.Stream") 
	objFSO.Type=1
	objFSO.Mode=admoderead
	objFSO.Open 
	objFSO.LoadFromFile(Path)
	if offset > 0 then 
		objFSO.Read(offset) 
	end if 
	if bytes = -1 then
		GetBytes =0
	else 
		GetBytes =Byte2Lng(objFSO.Read(bytes)) 
	end if 
	objFSO.Close 
	set objFSO = nothing 
end function 
Function Byte2Lng(bin)
  dim ret
  ret = 0
  for i = lenB(bin) to 1 step -1
   ret = ret *256 + ascb(midb(bin,i,1))
  next
  Byte2Lng=ret
End Function


FileName=Trim(Request.querystring("file"))
FileName="13.12.BMP"
n=instrRev( FileName ,".")
extPath=Lcase(mid(FileName ,n))
FileName=Server.MapPath(FileName)
If extPath=".bmp" Then
    Height = GetBytes(FileName,22,4) 
	Width = GetBytes(FileName,18,4)
	Response.write "ͼƬ���Ϊ:"
	Response.write  Width&"<BR>"
	Response.write "ͼƬ�߶�Ϊ��"
	Response.write  Height &"<BR>"
	Response.write  FileName&"��ͼƬ������ʾ��<BR>"
%>
<img src="<%=FileName%>" width=<%=Width%> height="<%=Height%>">
<%
Else
	Response.write "�޷�ȡ��ͼƬ��Ⱥ͸߶����ݣ�"
End If
%>
</body>

</html>
