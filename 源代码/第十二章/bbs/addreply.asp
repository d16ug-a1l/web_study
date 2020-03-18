
<%

rtid=Trim(Request.form("rtid"))
name=Trim(Request.form("name"))
pass=Trim(Request.form("pass"))
email=Trim(Request.form("email"))
url=Trim(Request.form("url"))
title=Trim(Request.form("title"))
nei=Trim(Request.form("nei"))
pic=Trim(Request.form("pic"))
face=Trim(Request.form("face"))
bl=0

	bad=split(badstr,"|")
	for i=0 to UBound(bad)
		nei=Replace(nei,bad(i),"**")
	next
	name=server.htmlencode(name)
		bad=split(badstr,"|")
	for i=0 to UBound(bad)
		name=Replace(name,bad(i),"**")
	next
	
	title=server.htmlencode(title)
		bad=split(badstr,"|")
	for i=0 to UBound(bad)
		title=Replace(title,bad(i),"**")
	next

  if len(title)>25 then
         title=left(title,25)
  end if
  if len(name)>5 then
         name=left(name,4)
  end if      
  


sql="select * from user where name='"&name&"' order by id desc"	  
set rs= Server.CreateObject("ADODB.RecordSet") 
rs.Open sql, Conn, 1, 2
if not rs.eof then
   if pass=rs("pass") and name=rs("name") then
        	  bl=1
	else
         	response.write"<SCRIPT language=JavaScript>alert('用户名和密码不匹配！');"
			response.write"javascript:history.go(-1)</SCRIPT>"
			response.end
       end if
else
    if pass<>"" then
		rs.addnew
    	rs("name")=name
    	rs("pass")=pass
    	rs("t1")=now()
		rs("logins")=1
    	rs.update
    	bl=1
	end if	
	rs.close
	set rs=nothing
end if

	

Set rs = Server.CreateObject("ADODB.Recordset")
rs.open "ly", conn,1,2 
     rs.addnew
	 rs("user_name")=name
	 rs("hfren")=name
     rs("title")=title
     rs("email")=email
     rs("url")=url
     rs("nei")=nei
     rs("t")=now()
	 rs("tt")=now()
     rs("pic")=pic
	 rs("face")=face
	 rs("lyip")=request.servervariables("remote_addr")
     rs("bl")=bl
     rs("hits")=1
     rs("jh")=0
	 rs("rt")=rtid
	 rs.update
	 
	 exec3="select * from ly where id=" + CStr(rtid) + " " 
     Set RS3 = Server.CreateObject("ADODB.RecordSet")                        
     rs3.Open exec3, Conn, 1, 2
     rs3("hfren")=name
	 rs3("tt")=now
	 y=rs3("hf")
     y=y+1
     rs3("hf")=y
     rs3.update
	 rs3.close
	 set rs3=nothing
%>
<meta http-equiv="refresh" content="3;url=show.asp?id=<%=rs("rt")%>">
<p><br>
</p>
<p> <br>
</p>
<table border="0" width="41%" cellspacing="1" cellpadding="3" bordercolorlight="#333333" bordercolordark="#FFFFFF" bgcolor="#205E7B" align="center">
  <tr>     
    <td  bgcolor="#205E7B" align="center" height="25"> <b>发 表 成 功</b> </td>
            </tr>
            <tr>
    <td width="100%" bgcolor="#92C8E2"style="line-height: 240%"> 
      <div align="center">系统将在3 秒后自动转到您所发表的帖子！</div>
        您也可以选择以下操作。<br>
        ·<a href="index.asp">返回论坛首页</a><br>
        ·<a href="show.asp?id=<%=rs("rt")%>">返回你所发表的帖子</a> 
    </td>
            </tr>
          </table>
		 <% rs.close
set rs=nothing
connclose()
%>
</body>
</html>