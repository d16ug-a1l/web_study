<!--#include file="md5.asp"-->
<%
  '如果尚未定义Pass对象，则将其定义为False，表示未登录
  If Request.Cookies("Order_Info") ="" Then
	'读取从表单传递过来的用户名和密码
    UserName = Request.Form("UserName")
	UserPwd = Request.Form("UserPwd")
	'用户名为空，显示错误信息
    If UserName = "" Then
      Errmsg = "请输入用户名和密码!"
    Else    
	  '连接数据库
      'Server对象的CreateObject方法建立Connection对象
      Set Conn=Server.CreateObject("ADODB.Connection")
      'Response.Write(Server.MapPath("User.mdb")&"<BR>")
	  Conn.ConnectionString="Provider=Microsoft.Jet.OLEDB.4.0;"&_
  			"Data Source="&Server.MapPath("STORE.mdb")
	  Conn.Open
	  Sql="select * from Users_Info where UserName='"&UserName &"' and  UserPwd='"&UserPwd &"'"
	   '读取用户数据
 
	  set rs=Conn.Execute(Sql) 
	  If rs.EOF Then  
         '用户不存在，显示错误信息
	    Errmsg = "用户不存在"
   	  Else      
           '登录成功
          Errmsg = ""
          Response.Cookies("Order_Info")("User")=MD5(UserName)
          Response.Cookies("Order_Info")("ID")=rs.Fields("UserID") 
    	  Response.Cookies("Order_Info").Expires=Date()+1
    	  Session("UserName") = UserName
	      Response.Redirect("index.asp")      
	  End If
    End If
 Else
 	 If IsEmpty(Session("UserName")) Then 
 	 	Response.Cookies("Order_Info")=""
 	Else
  	 	Response.Redirect("index.asp")
  	 End If
 End If
 '未登录或者登录不成功，显示登录界面
  If Request.Cookies("Order_Info")="" Then
%>
<HTML>
<HEAD><TITLE>请输入用户名和密码</TITLE></HEAD>
<BODY>
<p align="center"><font face="华文行楷" size="6" color="#0000FF">购&nbsp; 物&nbsp; 车&nbsp; 模&nbsp; 块</font></p> 

<p align="center"><font color="#800000">　<%=Errmsg%></font></p>
<form method="POST" action="logon.asp" name="Form" >
  <p align="center">用户名：&nbsp; <input type="text" name="UserName" size="20"></p>
  <p align="center">密&nbsp; 码：&nbsp; <input type="password" name="UserPwd" size="20"></p>
  <p align="center"><input type="submit" value="提交" name="B1"><input type="reset" value="全部重写" name="B2"></p>
</form>
<p align="center">　</p>
</BODY>
</HTML>
<%
    Response.End
  End If
%>