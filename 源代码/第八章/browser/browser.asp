<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>���������</title>
</head>

<body>
<table align=center border=1>
<tr><td>
	<p align="center"><font face="�����п�" color="#0000FF" size="6">�������Ϣ</font></td></tr>
<%
Dim Obj
Set Obj=Server.CreateObject("MSWC.BrowserType")
Response.Write(Request.ServerVariables("HTTP_USER_AGENT")&"<BR>")
Response.Write("<TR><TD>")
Response.Write("���������ƣ�"&Obj.Browser&"<BR>")
Response.Write("</TD></TR><TR><TD>")
Response.Write("������汾��"&Obj.Version&"<BR>")
Response.Write("</TD></TR><TR><TD>")
Response.Write("���汾�ţ�"&Obj.Majorver&"<BR>")
Response.Write("</TD></TR><TR><TD>")
Response.Write("�ΰ汾�ţ�"&Obj.Minorver&"<BR>")
Response.Write("</TD></TR><TR><TD>")
Response.Write("������ƽ̨��"&Obj.Platform&"<BR>")
Response.Write("</TD></TR><TR><TD>")
Response.Write("�������Ƿ�֧��Cookies��"&Obj.Cookies&"<BR>")
Response.Write("</TD></TR><TR><TD>")
Response.Write("�������Ƿ�֧�ֿ�ܣ�"&Obj.Frames&"<BR>")
Response.Write("</TD></TR><TR><TD>")
Response.Write("�������Ƿ�֧��JavaScript��"&Obj.JavaScript&"<BR>")
Response.Write("</TD></TR><TR><TD>")
Response.Write("�������Ƿ�֧��JavaApplets:"&Obj.JavaApplets&"<BR>")
Response.Write("</TD></TR><TR><TD>")
Response.Write("�������Ƿ�֧�ֱ��"&Obj.Tables&"<BR>")
Response.Write("</TD></TR><TR><TD>")
Response.Write("�������Ƿ�֧��VBScript��"&Obj.VBScript&"<BR>")
Response.Write("</TD></TR><TR><TD>")
Response.Write("�������Ƿ�֧�ֱ���������"&Obj.BackGroundSounds&"<BR>")
Response.Write("</TD></TR><TR><TD>")
Response.Write("�������Ƿ�֧��ActiveControls��"&Obj.ActiveXControls&"<BR>")
Response.Write("</TD></TR>")
%>
</table>
</body>

</html>