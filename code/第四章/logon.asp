<html>
<head>
<title>处理提交的数据</title>
</head>
<body>
<P><form >
<!---设置字体的颜色为红色--->
 <font color="red">
    <!---使用request.form获取名为”user”文本框的值--->
	<%=request.form("user")%>
</font>
 登录成功！</P>
<p>你输入的密码是:
<!---设置字体的颜色为红色--->
<font color="red">
<!---使用request.form获取名为” password”文本框的值--->
	<%=request.form("password")%>
</font></P>
</form>
</body>
</html>
