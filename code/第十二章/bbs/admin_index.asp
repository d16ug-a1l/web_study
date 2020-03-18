<%dim action
action=request.QueryString("action")
action=replace(action,"'","")
%>

<HTML>
<HEAD>
<TITLE>论坛登陆</TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=gb2312">
<style type="text/css">
<!--
body {
	margin: 0;
	overflow: hidden;
	scrollbar-face-color: D9E5F6;
	scrollbar-highlight-color: #FFFFFF;
	scrollbar-shadow-color: darkseablue;
	scrollbar-3dlight-color: D9E5F6;
	scrollbar-arrow-color: darkseablue;
	scrollbar-track-color: #f3faf4;
	scrollbar-darkshadow-color: #f3faf4;
}
td {
	font-size: 12px;
	line-height: 140%;
}
.copyright {
	padding-bottom: 10px;
}
.sysname {
	padding-bottom: 5px;
}
a:link {
	color: #000000;
	text-decoration: none;
}
a:visited {
	color: #000000;
	text-decoration: none;
}
a:hover {
	color: red;
	text-decoration: underline;
}
-->
</style>
</HEAD>
<body  text="#000000" leftmargin="0" topmargin="0" oncontextmenu=""return false;"">

<%
if  session("user_name")<>"" and session("bz")<>""  and session("bz")<>"0" then
%>
<table width="95%" border=0 cellspacing=1 cellpadding=0 align=center>
<tr><td height=50>
<a href="index.asp">返回首页</a>&nbsp;&nbsp;<a href="admin_index.asp?action=myuser">用户管理</a>&nbsp;&nbsp;<a href="admin_index.asp?action=tz">帖子管理</a>&nbsp;&nbsp;<a href="admin_index.asp?action=xzb">公告管理</a>&nbsp;&nbsp;<a href="login.asp?action=out">退出管理</a>
</td></tr></table>


      <%
end if
select case action
case ""	%>


<%
if  session("user_name")="" or session("bz")<>"2" then
%>


      <% response.write"<table width=""100%"" border=0 cellspacing=1 cellpadding=0 height=50><tr><td height=100></td></tr></table>"
      response.write"<table width=""50%"" border=""0"" cellspacing=""1"" cellpadding=""3"" align=""center"" bgcolor=#205E7B>"
      response.write" <form name=""form1"" method=""post"" action=""login.asp""><tr>"
            response.write"<td colspan=""2"" height=""25"" align=""center""><font color=""#FFFFFF"">管理登陆</font></td></tr><tr>"
			response.write"<td bgcolor=#64B3D9 align=""center"" width=""40%"">帐号： </td><td bgcolor=#64B3D9 width=""60%"">" 
              response.write"<input type=""text"" onMouseOver=""this.style.backgroundColor =#E5F0FF""title=""输入管理员名称"" style=""BORDER-RIGHT: #b4b4b4 1px double; BORDER-TOP: #b4b4b4 1px double; BORDER-LEFT: #b4b4b4 1px double; COLOR: #ff90cd; BORDER-BOTTOM: #b4b4b4 1px double; BACKGROUND-COLOR: #ffffff"" onMouseOut=""this.style.backgroundColor = ''"" name=""name"" size=""16"" value=>"
          response.write"</td></tr><tr><td bgcolor=#64B3D9 align=""center"" width=""40%"">密码： </td>"
            response.write"<td bgcolor=""#64B3D9"" width=""60%""> "
              response.write"<input onMouseOver=""this.style.backgroundColor = '#E5F0FF'"" title='输入管理员密码' style=""BORDER-RIGHT: #b4b4b4 1px double; BORDER-TOP: #b4b4b4 1px double; BORDER-LEFT: #b4b4b4 1px double; COLOR: #ff90cd; BORDER-BOTTOM: #b4b4b4 1px double; BACKGROUND-COLOR: #ffffff"" onMouseOut=""this.style.backgroundColor = ''"" name=pass size=""16"" maxlength=""15"" type=""password"">"
          response.write"</td></tr><tr><td colspan=""2"" align=""center"">" 
             response.write" <input onMouseOver=""this.style.backgroundColor='#FFC864'"" style=""BORDER-RIGHT: 0px solid; BORDER-TOP: 0px solid; BORDER-LEFT: 0px solid; COLOR: #000000; BORDER-BOTTOM: 0px solid; BACKGROUND-COLOR: #3399CC"" onMouseOut=""this.style.backgroundColor='#3399CC'"" type=submit value=""登 陆"" name=""submit""> "
            response.write" <input onMouseOver=""this.style.backgroundColor='#FFC864'"" style=""BORDER-RIGHT: 0px solid; BORDER-TOP: 0px solid; BORDER-LEFT: 0px solid; COLOR: #000000; BORDER-BOTTOM: 0px solid; BACKGROUND-COLOR: #3399CC"" onMouseOut=""this.style.backgroundColor='#3399CC'"" type=reset value=""取 消"" name=""submit2"">"
          response.write"</td></tr></form></table>"%>

<%
else
%>
<table width="90%" border=0 cellspacing=1 cellpadding=0 align=center>
<tr><td height=20 align=center><font color=blue>
管理员已成功登陆，管理结束后请单击“退出管理”。部分管理功能需要在前台进行（例如删除跟帖），请单击“返回首页”，浏览帖子时可删除跟帖。</font>
</td></tr></table>


<%
end if
%>

      <%case "myuser"	%>
      <br>
      <table width="96%" border="0" align="center" cellspacing="1" bgcolor="#205E7B">
        <tr align="center"> 
          <td  height="25"><font color="#FFFFFF">ID</font></td>
          <td  height="25"><font color="#FFFFFF">会员帐号（单击修改密码）</font></td>
 
          <td  height="25"><font color="#FFFFFF">级别</font></td>
          <td height="25" colspan="4"><font color="#FFFFFF">操作</font></td>
        </tr>
        <tr> 
          <td height="1" colspan="9" bgcolor="#000000"></td>
        </tr>
        <%
			 sql="select * from user order by id desc" 
        set rs=server.createobject("adodb.recordset")
		rs.open sql,conn,3,3
        if rs.eof then

%>
        <tr align="center"> 
          <td height="25" colspan="9" bgcolor="#f0f0f0">没有用户注册</td>
        </tr>
        <%else
		if not isempty(request("page")) then   
		pagecount=cint(request("page"))  
				else
		pagecount=1
		end if

		rs.pagesize=10
		rs.AbsolutePage=pagecount	  
        do while not rs.eof%>
        <tr bgcolor="#f0f0f0"> 
          <td width="52" height="25" align="center"><%=rs("id")%></td>
          <td width="116" height="25"><a href="admin_index.asp?action=admin&id=<%=rs("id")%>&name=<%=rs("name")%>"><font color="#0000FF"><%=rs("name")%></font></td>
          
          <td width="132" height="25"> 
            <div align="center"> 
              <%
if rs("bz")=2 then 
 response.write"管理员" 
else
 response.write"会员" 
end if%>
            </div>
          </td>
          <td width="83" height="25" align="center"><a href="guanli.asp?action=myuser&action1=hy&id=<%=rs("id")%>&page=<%=cstr(pagecount)%>"><font color="#0000FF">降为会员</font></a></td>
          <td width="83" height="25" align="center"><a href="guanli.asp?action=myuser&action1=zbz&id=<%=rs("id")%>&page=<%=cstr(pagecount)%>"><font color="#0000FF">设为管理员</font></a></td>
          <td width="84" height="25" align="center"><a href="guanli.asp?action=myuser&action1=del&id=<%=rs("id")%>&page=<%=cstr(pagecount)%>"><font color="#FF0000">删除</font></a></td>
        </tr>
        <%
rs.movenext
i=i+1                                                                     
if i>=rs.pagesize then exit do                                                           
loop
%>
        <tr> 
          <form action="admin_index.asp?action=myuser" method="post">
            <td height="35" colspan="13" bgcolor="#D6DFF7"> 
              <div align="center"> 共 <b><%=rs.recordcount%></b> 位用户, 页次: <b><font color=red><%=pagecount%></font>/<%=rs.pagecount%></b>, 
                当前从第 
                <%
if pagecount<=1 then
response.write "<font color=red>1</font>"
else
response.write "<font color=red>" & pagecount*rs.pagesize-rs.pagesize+1 & "</font>"
end if
%>
                位开始。 
                <% if pagecount=1 and rs.pagecount<>pagecount and rs.pagecount<>0 then%>
                <a href="admin_index.asp?id=<%=id%>&action=myuser&page=<%=cstr(pagecount+1)%>">下一页</a> 
                <% end if %>
                <% if rs.pagecount>1 and rs.pagecount=pagecount then %>
                <a href="admin_index.asp?id=<%=id%>&action=myuser&page=<%=cstr(pagecount-1)%>"> 
                上一页</a> 
                <%end if%>
                <% if pagecount<>1 and rs.pagecount<>pagecount then%>
                <a href="admin_index.asp?id=<%=id%>&action=myuser&page=<%=cstr(pagecount-1)%>"> 
                上一页</a> <a href="admin_index.asp?id=<%=id%>&action=myuser&page=<%=cstr(pagecount+1)%>"> 
                下一页</a> 
                <% end if%>
                &nbsp; 直接到第 
                <select name="page">
                  <%for i=1 to rs.pagecount%>
                  <option value="<%=i%>"><%=i%></option>
                  <%next%>
                </select>
                页 
                <input type="submit" name="go" value="Go">
                <input type="hidden" name="id" value=<%=id%>>
              </div>
            </td>
          </form>
        </tr>
        <%
end if
rs.close
set rs=nothing
%>
      </table>
	  <%case "tz"%><br>
      <table width="96%" border="0" cellspacing="1" cellpadding="3" align="center" bgcolor="#205E7B">
        <tr align="center"> 
        <td width="10%" height="25"><font color="#FFFFFF">ID</font></td>
	<td width="40%" height="25"><font color="#FFFFFF">帖子标题</font></td>
	<td width="15%"><font color="#FFFFFF">发帖时间</font></td>
	<td width="10%"><font color="#FFFFFF">帖子属性</font></td>
	<td width="25%"><font color="#FFFFFF">操作</font></td>
        </tr>
        <tr> 
          <td colspan="4" height="1" bgcolor="#000000"></td>
        </tr>
        <%
sql="select * from ly where rt=0 order by gd desc,id desc"	
set rs=server.createobject("ADODB.Recordset")
rs.open sql,conn,1,1
if rs.eof then
response.write"<tr><td height=25  colspan='4' bgcolor=#f0f0f0 align=center>没有帖子</td></tr>"
else
if not isempty(request("page")) then   
		pagecount=cint(request("page"))  
else
		pagecount=1
end if

		rs.pagesize=12
		rs.AbsolutePage=pagecount	  
        do while not rs.eof
%>
        <tr> 
            <td bgcolor="#f0f0f0"  height="25" align="center"><%=rs("id")%></td>
            
          <td bgcolor="#f0f0f0"  height="25"><a href="show.asp?id=<%=rs("id")%>" target="_blank"><%=rs("title")%></a>(<font color="#0000FF">回帖数：<%=rs("hf")%></font>)</td>
          <td bgcolor="#f0f0f0" align=center><%=rs("t")%></td>
          <td bgcolor="#f0f0f0" align="center"><font color=blue>
<%
if rs("jh")=0 then response.write"[普通]"
if rs("jh")=1 then response.write"[精华]"
if rs("jh")=2 then response.write"[锁定]"
if rs("gd")=1 then response.write"[固定]"
%></font>
<td bgcolor="#f0f0f0" align="center">
<a href="guanli.asp?action1=gd&id=<%=rs("id")%>">固顶</a>&nbsp;
<a href="guanli.asp?action1=jg&id=<%=rs("id")%>">解固</a>&nbsp;
<a href="guanli.asp?action1=jh&id=<%=rs("id")%>">精华</a>&nbsp;
<a href="guanli.asp?action1=sd&id=<%=rs("id")%>">锁定</a>&nbsp;
<a href="guanli.asp?action1=js&id=<%=rs("id")%>">普通</a>&nbsp;
<a href="guanli.asp?action1=deltz&id=<%=rs("id")%>">删除</a>
</td>
        </tr>
        <%
rs.movenext
i=i+1                                                                     
if i>=rs.pagesize then exit do                                                           
loop
%>
		  <tr bgcolor="#92C8E2"> 
          <form action="admin_index.asp?action=tz" method="post">
            <td height="35" colspan="11"> 
              <div align="center"> 共 <b><%=rs.recordcount%></b> 帖子, 页次: <b><font color=red><%=pagecount%></font>/<%=rs.pagecount%></b>, 
                当前从第 
                <%
if pagecount<=1 then
response.write "<font color=red>1</font>"
else
response.write "<font color=red>" & pagecount*rs.pagesize-rs.pagesize+1 & "</font>"
end if
%>
                位开始。 
                <% if pagecount=1 and rs.pagecount<>pagecount and rs.pagecount<>0 then%>
                <a href="admin_index.asp?action=tz&page=<%=cstr(pagecount+1)%>">下一页</a> 
                <% end if %>
                <% if rs.pagecount>1 and rs.pagecount=pagecount then %>
                <a href="admin_index.asp?action=tz&page=<%=cstr(pagecount-1)%>"> 
                上一页</a> 
                <%end if%>
                <% if pagecount<>1 and rs.pagecount<>pagecount then%>
                <a href="admin_index.asp?action=tz&page=<%=cstr(pagecount-1)%>"> 
                上一页</a> <a href="admin_index.asp?action=tz&page=<%=cstr(pagecount+1)%>"> 
                下一页</a> 
                <% end if%>
                &nbsp; 直接到第 
                <select name="page">
                  <%for i=1 to rs.pagecount%>
                  <option value="<%=i%>"><%=i%></option>
                  <%next%>
                </select>
                页 
                <input type="submit" name="go" value="Go">
                <input type="hidden" name="id" value=<%=id%>>
              </div>
            </td>
          </form>
        </tr>
        <%
end if
rs.close
set rs=nothing
%>
      </table>
      <%case "xzb"%><br>
      <table width="96%" border="0"" cellspacing="1" cellpadding="3" align="center" bgcolor="#205E7B">
        <tr>
           <td width=""%" height="25" align="center"><font color="#FFFFFF">ID</font></td>
            <td width="64%" height="25" align="center"><font color="#FFFFFF">标题</font></td>
           <td width=""9%" align="center"><font color="#FFFFFF">时间</font></td>
            <td width="9%" align="center"><font color="#FFFFFF">操作</font></td></tr>
       <tr><td colspan="4" height="1" bgcolor="#000000" align="center"></td></tr>
       
<%sql="select * from smallpager order by id desc"	
set rs=server.createobject("ADODB.Recordset")
rs.open sql,conn,1,1
if rs.eof then
response.write"<tr><td height=25  colspan='4' bgcolor=#f0f0f0 align=center>没有公告</td></tr>"
else
if not isempty(request("page")) then   
		pagecount=cint(request("page"))  
else
		pagecount=1
end if

		rs.pagesize=12
		rs.AbsolutePage=pagecount	  
        do while not rs.eof%>
        <tr>
			<td bgcolor="#f0f0f0" width="8%" height="25" align="center"><%=rs("id")%></td>
            <td bgcolor="#f0f0f0" width="64%" height="25"><%=rs("s_title")%></td>
            <td bgcolor="#f0f0f0" width="19%" align="center"><%=rs("adate")%></td>
            <td bgcolor="#f0f0f0" width="9%" align="center"><a href="guanli.asp?action=xzb&action1=delxzb&id=<%=rs("id")%>&page=<%=cstr(pagecount)%>"><font color="#FF0000">删除</font></a> 
            </td>
        </tr>
        <%
rs.movenext
i=i+1                                                                     
if i>=rs.pagesize then exit do                                                           
loop
%>
		  <tr bgcolor="#92C8E2" align="right">
          <form action="admin_index.asp?action=xzb" method="post">
            <td height="35" colspan="11" bgcolor="#D6DFF7" align="center"> 
              <div align="center"> 共 <b><%=rs.recordcount%></b> 帖子, 页次: <b><font color=red><%=pagecount%></font>/<%=rs.pagecount%></b>, 
                当前从第 
                <%
if pagecount<=1 then
response.write "<font color=red>1</font>"
else
response.write "<font color=red>" & pagecount*rs.pagesize-rs.pagesize+1 & "</font>"
end if
%>
                位开始。 
                <% if pagecount=1 and rs.pagecount<>pagecount and rs.pagecount<>0 then%>
                <a href="admin_index.asp?id=<%=id%>&action=myuser&page=<%=cstr(pagecount+1)%>">下一页</a> 
                <% end if %>
                <% if rs.pagecount>1 and rs.pagecount=pagecount then %>
                <a href="admin_index.asp?id=<%=id%>&action=myuser&page=<%=cstr(pagecount-1)%>"> 
                上一页</a> 
                <%end if%>
                <% if pagecount<>1 and rs.pagecount<>pagecount then%>
                <a href="admin_index.asp?id=<%=id%>&action=myuser&page=<%=cstr(pagecount-1)%>"> 
                上一页</a> <a href="admin_index.asp?id=<%=id%>&action=myuser&page=<%=cstr(pagecount+1)%>"> 
                下一页</a> 
                <% end if%>
                &nbsp; 直接到第 
                <select name="page">
                  <%for i=1 to rs.pagecount%>
                  <option value="<%=i%>"><%=i%></option>
                  <%next%>
                </select>
                页 
                <input type="submit" name="go" value="Go">
                <input type="hidden" name="id" value=<%=id%>>
              </div>
            </td>
          </form>
        </tr>
        <%
end if
rs.close
set rs=nothing
%>
      </table>




 <%case "admin"%>

<%
if request("ok") ="ok" and request("pass")=request("pass1") then
conn.execute "update user set pass='"&request("pass")&"' where id=" & request("id")

			response.write "<script language='javascript'>"
			response.write "alert('会员密码修改成功！');"
			response.write "location.href='admin_index.asp?action=myuser';"			
			response.write "</script>"
else
%>

<form action="admin_index.asp?action=admin&name=<%=request("name")%>" method="post">
      <table width="96%" border="0"" cellspacing="1" cellpadding="3" align="center">
        <tr><td>修改会员密码</td></tr>
	<tr><td>会员名称：<%=request("name")%></td></tr>
	<tr><td>修改密码：<input type="text" name="pass" value=""></td></tr>
	<tr><td>重复密码：<input type="text" name="pass1" value=""></td></tr>
	<tr><td><input type="submit" name="admin" value="确认"></td></tr>

 <input type="hidden" name="id" value="<%=request("id")%>">
<input type="hidden" name="ok" value="ok">

</table>
</form>


<%end if%>





      <%end select%>


</TABLE>
</BODY>
</HTML>