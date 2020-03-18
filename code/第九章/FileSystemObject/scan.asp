
<%
FileName=Request.querystring("file")
Set objfilesys=createobject("scripting.filesystemobject")
FilePath=Server.MapPath(FileName)
set objfile=objfilesys.opentextfile(FilePath,1 )
full=objfile.readall
response.write Server.HTMLEncode(full)
%>
