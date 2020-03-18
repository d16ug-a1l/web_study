<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>浏览器类型</title>
</head>

<body>
<table align=center border=1>
<tr><td>
	<p align="center"><font face="华文行楷" color="#0000FF" size="6">浏览器信息</font></td></tr>
<%
Dim Obj
Set Obj=Server.CreateObject("MSWC.BrowserType")
Response.Write(Request.ServerVariables("HTTP_USER_AGENT")&"<BR>")
Response.Write("<TR><TD>")
Response.Write("服务器名称："&Obj.Browser&"<BR>")
Response.Write("</TD></TR><TR><TD>")
Response.Write("浏览器版本："&Obj.Version&"<BR>")
Response.Write("</TD></TR><TR><TD>")
Response.Write("主版本号："&Obj.Majorver&"<BR>")
Response.Write("</TD></TR><TR><TD>")
Response.Write("次版本号："&Obj.Minorver&"<BR>")
Response.Write("</TD></TR><TR><TD>")
Response.Write("服务器平台："&Obj.Platform&"<BR>")
Response.Write("</TD></TR><TR><TD>")
Response.Write("服务器是否支持Cookies："&Obj.Cookies&"<BR>")
Response.Write("</TD></TR><TR><TD>")
Response.Write("服务器是否支持框架："&Obj.Frames&"<BR>")
Response.Write("</TD></TR><TR><TD>")
Response.Write("服务器是否支持JavaScript："&Obj.JavaScript&"<BR>")
Response.Write("</TD></TR><TR><TD>")
Response.Write("服务器是否支持JavaApplets:"&Obj.JavaApplets&"<BR>")
Response.Write("</TD></TR><TR><TD>")
Response.Write("服务器是否支持表格："&Obj.Tables&"<BR>")
Response.Write("</TD></TR><TR><TD>")
Response.Write("服务器是否支持VBScript："&Obj.VBScript&"<BR>")
Response.Write("</TD></TR><TR><TD>")
Response.Write("服务器是否支持背景声音："&Obj.BackGroundSounds&"<BR>")
Response.Write("</TD></TR><TR><TD>")
Response.Write("服务器是否支持ActiveControls："&Obj.ActiveXControls&"<BR>")
Response.Write("</TD></TR>")
%>
</table>
</body>

</html>