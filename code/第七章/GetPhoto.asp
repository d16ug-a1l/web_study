
<%
ID=Request("FileName")
Call DisplayBMP(4)
Sub  DisplayBMP(ID)
	ID=Cint(ID)
	Set Conn= Server.CreateObject("ADODB.Connection")
	Conn.ConnectionString="Provider=Microsoft.Jet.OLEDB.4.0;"&_
				 			"Data Source="&Server.MapPath("img.mdb")
	Conn.Open
	Set rs=Conn.Execute("select * from [image] where id="&ID)
	FilePath=rs("FilePath")
	n=instrRev(FilePath,".")
	FSize = rs("img").ActualSize
    Response.Buffer = true
	extPath=Ucase(mid(FilePath,n+1))
	Response.Expires = -1
	Response.ContentType = "image/"&extPath
	Response.BinaryWrite rs("img").getChunk(FSize )
	Response.end
End Sub 
%>