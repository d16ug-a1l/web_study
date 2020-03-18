<!--#include file="funciton.asp"-->
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>新建网页 1</title>
</head>

<body>
<p align="center"><font face="华文行楷" size="6" color="#0000FF">权 限 管 理 模 块</font></p> 
<%

typea=Request.QueryString("type")
Response.write(typea)
ID=Request.QueryString("ID")
Response.write(ID)
If ID=0 Then 
	Response.End
End If



%>
<p><input type="submit" value="确定" name="B1"><input type="reset" value="重置" name="B2"></p>	
</form>
</body>

</html>