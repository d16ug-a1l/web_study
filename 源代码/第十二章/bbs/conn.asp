<%@LANGUAGE="VBSCRIPT"%>
<%	dim conn
	dim connstr
	db="data/forum.mdb"
	Set conn = Server.CreateObject("ADODB.Connection")
	connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(DB)
'如果你的服务器采用较老版本Access驱动，请用下面连接方法
	connstr="driver={Microsoft Access Driver (*.mdb)};dbq=" & Server.MapPath(DB)
	conn.Open connstr
	
Sub connclose()
    conn.close()
	set conn=nothing
End Sub

badstr=""

footer="本课程论坛无需注册即可发帖。论坛言论仅代表发帖者个人看法，不代表本站观点。"
sitename="民商法"
siteurl=""
	
%>
