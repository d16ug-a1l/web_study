<%@LANGUAGE="VBSCRIPT"%>
<%	dim conn
	dim connstr
	db="data/forum.mdb"
	Set conn = Server.CreateObject("ADODB.Connection")
	connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(DB)
'�����ķ��������ý��ϰ汾Access�����������������ӷ���
	connstr="driver={Microsoft Access Driver (*.mdb)};dbq=" & Server.MapPath(DB)
	conn.Open connstr
	
Sub connclose()
    conn.close()
	set conn=nothing
End Sub

badstr=""

footer="���γ���̳����ע�ἴ�ɷ�������̳���۽��������߸��˿�����������վ�۵㡣"
sitename="���̷�"
siteurl=""
	
%>
