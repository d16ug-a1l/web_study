<html>

<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>在字段搜索</title>
</head>

<body>

<form method="POST" action="Search.asp">
	<p align="center"><font face="华文行楷" size="6" color="#0000FF">搜 索 模 块</font></p>
	<p align="center"><font face="华文行楷" color="#0000FF">在字段</font><select size="1" name="Field">
	<%
	  Session("prev_search")=""
	  Set Conn=Server.CreateObject("ADODB.Connection")
      'Response.Write(Server.MapPath("User.mdb")&"<BR>")
	  Conn.ConnectionString="Provider=Microsoft.Jet.OLEDB.4.0;"&_
  			"Data Source="&Server.MapPath("book.mdb")
	  Conn.Open
	  Sql="select * from Book_Info"
	  set rs=Conn.Execute(Sql) 
	  If rs.EOF=false Then  
         For i=0 To rs.Fields.Count-1
         	Response.write("<option value='"&rs.Fields(i).Name&","&rs.Fields(i).type&"'>"&rs.Fields(i).Name&"</option>")
         Next   	  	 
   	  End If
	%>
	</select><font face="华文行楷" color="#0000FF">搜索</font>
	<input type="text" name="Content" size="16">
	<select size="1" name="Type">
	<option selected value="MH">模糊</option>
	<option value="JQ">精确</option>
	</select><input type="submit" value="搜索" name="B1"></p>
	<p align="center"><input type="radio" value="Reset" checked name="R1">重新搜索<input type="radio" name="R1" value="Result">在结果中查询</p>
</form>
 
</body>

</html>