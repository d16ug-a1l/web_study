<%
FileName=Request("FileName")
ReferWEB=trim(request.ServerVariables("http_referer"))
WEB=trim(request.ServerVariables("server_name"))

Dim n
n=Instr(8,ReferWEB,"/")
ReferWEB=mid(ReferWEB,8,n-8)
If ReferWEB=WEB Then
	Call DisplayBMP(FileName)
Else
	Response.write "µÁÁ´"
End If

Sub  DisplayBMP(FileName)
n=instrRev(FileName,".")
extPath=Ucase(mid(FileName,n))

If extPath=".BMP" or extPath=".JPG" or extPath=".GIF" Then
	FileName=Server.MapPath(FileName)
	Set objFSO = CreateObject("ADODB.Stream") 
	objFSO.Type=1
	objFSO.Mode=admoderead
	objFSO.Open 
	objFSO.LoadFromFile(FileName)
	Response.Expires = -1
	Response.AddHeader "Pragma","no-cache"  
	Response.AddHeader "cache-ctrol","no-cache" 
	Response.ContentType = "Image/"& extPath
	Response.BinaryWrite objFSO.Read
End If
End Sub 
%>
