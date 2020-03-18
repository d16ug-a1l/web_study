<title>在结果中再搜索</title>
<body bgcolor="#FFFFFF">
<%
u_search=request.form("u_search")
u_prev_search=request.form("u_prev_search")
u_search_within=request.form("u_search_within")


if u_search <> "" then

    if u_prev_search = "" then 
        u_prev_search=u_search
else

    u_prev_search=u_prev_search &","& u_search
    g_prev_search=split(u_prev_search,",")
    num_inputted=ubound(g_prev_search)
end if

sql= "select * from states where (capital like '%%"& u_search & "%%') "

if u_search_within = "Yes" then
    for counter =0 to num_inputted-1
        sql=sql& "and (capital like '%%"& g_prev_search(counter) & "%%') "
    next
end if 

accessdb="state_info" 
cn="DRIVER={Microsoft Access Driver (*.mdb)};"
cn=cn & "DBQ=" & server.mappath(accessdb)

Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open sql, cn
' 如果没有找到相应的信息
if rs.eof then
%>
    没有任何记录 

<%' 有相应的信息就列出来
else 
    rs.movefirst
    do while Not rs.eof
%>
        <%= rs("capital") %><br>
<%
    rs.movenext
    loop
end if 
end if 
%>
<!-- Begin Form Input Area -->
<form action="<%= request.servervariables("script_name") %>" method="post">
<input type="text" name="u_search" value="<%= u_search %>">
<br>
<%
if u_search <> "" then %>
<input type = "radio" name="u_search_within" checked value="No"> 重新搜索 &nbsp; 
<input type = "radio" name="u_search_within" value="Yes"> 在结果中搜索 
<%
if u_search_within = "Yes" then %>
<input type = "hidden" name="u_prev_search" value="<%= u_prev_search %>">
<%
else %>
<input type = "hidden" name="u_prev_search" value="<%= u_search %>">
<% end if%>
<br>
<% end if%>
<input type="submit" value="搜索">
</form>
<!-- End Form Input Area -->
<p>　</p>
<%= sql %>