<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>ɾ��Contents</title>
</head>
<body>
<form method="POST" action="DeleContents.asp">
<SELECT NAME="ContentRemove" SIZE="1">
<%
'��ȡContents�����е�����Ԫ�ز�����Select��
For Each objItem in Application.Contents
	Response.Write "<OPTION value='"&objItem &"'>" & objItem & "</OPTION>"
Next
%>
</select>
<input type="checkbox" name="C1" value="ON">ȫ��ɾ��<p>
<input type="submit" value="�ύ" name="B1"><input type="reset" value="����" name="B2"></p>
</p>
</form>
</body>
</html>
