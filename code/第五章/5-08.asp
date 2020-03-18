<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>删除Contents</title>
</head>
<body>
<form method="POST" action="DeleContents.asp">
<SELECT NAME="ContentRemove" SIZE="1">
<%
'获取Contents集合中的所有元素并放入Select中
For Each objItem in Application.Contents
	Response.Write "<OPTION value='"&objItem &"'>" & objItem & "</OPTION>"
Next
%>
</select>
<input type="checkbox" name="C1" value="ON">全部删除<p>
<input type="submit" value="提交" name="B1"><input type="reset" value="重置" name="B2"></p>
</p>
</form>
</body>
</html>
