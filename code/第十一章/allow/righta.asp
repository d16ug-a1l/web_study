<!--#include file="funciton.asp"-->
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�½���ҳ 1</title>
</head>

<body>
<p align="center"><font face="�����п�" size="6" color="#0000FF">Ȩ �� �� �� ģ ��</font></p> 
<%

typea=Request.QueryString("type")
Response.write(typea)
ID=Request.QueryString("ID")
Response.write(ID)
If ID=0 Then 
	Response.End
End If



%>
<p><input type="submit" value="ȷ��" name="B1"><input type="reset" value="����" name="B2"></p>	
</form>
</body>

</html>