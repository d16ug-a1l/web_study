<%if session("bz")<>"2" then response.end%>
<%
action1=trim(request("action1"))
action1=replace(action1,"'","")
action=trim(request("action"))
action=replace(action,"'","")
id=trim(request("id"))
id=replace(id,"'","")
page=trim(request("page"))
page=replace(page,"'","")

if request("action1")="del" and request("id")<>"" then
conn.execute "delete * from user where id=" & request("id")

response.redirect "admin_index.asp?action=myuser"
response.end
end if

if request("action1")="hy" and request("id")<>"" then

conn.execute "update user set bz=0 where id=" & request("id")

response.redirect "admin_index.asp?action=myuser"
response.end
end if


if request("action1")="zbz" and request("id")<>"" then

conn.execute "update user set bz=2 where id=" & request("id")

response.redirect "admin_index.asp?action=myuser"
response.end
end if


if request("action1")="delxzb" and request("id")<>"" then

conn.execute "delete * from smallpager where id=" & request("id")

response.redirect "admin_index.asp?action=xzb"
response.end
end if

if request("action1")="gd" and request("id")<>"" then

conn.execute "update ly set gd=1 where id=" & request("id")

response.redirect "admin_index.asp?action=tz"
response.end
end if


if request("action")="gd" and request("id")<>"" then

conn.execute "update ly set gd=1 where id=" & request("id")

response.redirect "show.asp?id=" & request("id")
response.end
end if

if request("action1")="jg" and request("id")<>"" then

conn.execute "update ly set gd=0 where id=" & request("id")

response.redirect "admin_index.asp?action=tz"
response.end
end if


if request("action")="jg" and request("id")<>"" then

conn.execute "update ly set gd=0 where id=" & request("id")

response.redirect "show.asp?id=" & request("id")
response.end
end if



if request("action1")="deltz" and request("id")<>"" then

conn.execute "delete * from ly where rt=" & request("id")
conn.execute "delete * from ly where id=" & request("id")

response.redirect "admin_index.asp?action=tz"
response.end
end if


if request("action")="deltz" and request("id")<>"" then

conn.execute "delete * from ly where rt=" & request("id")
conn.execute "delete * from ly where id=" & request("id")

response.redirect "index.asp"
response.end
end if


if request("action")="delgt" and request("id")<>"" and request("id1")<>"" then

conn.execute "delete * from ly where id=" & request("id1")

response.redirect "show.asp?id=" &  request("id")
response.end
end if


if request("action")="jh" and request("id")<>"" then

conn.execute "update ly set jh=1 where id=" & request("id")

response.redirect "show.asp?id=" & request("id")
response.end
end if

if request("action1")="jh" and request("id")<>"" then

conn.execute "update ly set jh=1 where id=" & request("id")

response.redirect "admin_index.asp?action=tz"
response.end
end if


if request("action")="sd" and request("id")<>"" then

conn.execute "update ly set jh=2 where id=" & request("id")

response.redirect "index.asp"
response.end
end if

if request("action1")="sd" and request("id")<>"" then

conn.execute "update ly set jh=2 where id=" & request("id")

response.redirect "admin_index.asp?action=tz"
response.end
end if



if request("action")="js" and request("id")<>"" then

conn.execute "update ly set jh=0 where id=" & request("id")

response.redirect "index.asp"
response.end
end if


if request("action1")="js" and request("id")<>"" then

conn.execute "update ly set jh=0 where id=" & request("id")

response.redirect "admin_index.asp?action=tz"
response.end
end if


response.end

%>
