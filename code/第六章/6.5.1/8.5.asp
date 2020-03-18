<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>新建网页 1</title>
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
	Response.write "图片宽度为:"
	Response.write  Width&"<BR>"
	Response.write "图片高度为："
	Response.write  Height &"<BR>"
	Response.write  FileName&"的图片如下所示。<BR>"
%>
<img src="<%=FileName%>" width=<%=Width%> height="<%=Height%>">
<%
Else
	Response.write "无法取得图片宽度和高度数据！"
End If
%>
</body>

</html>
