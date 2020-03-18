
javastr="<style>A.menu2{text-decoration: none;color:#000099;line-height: 25pt;font-size:9pt} A.menu2:hover {text-decoration: none;line-height: 25pt;font-size:9pt;color: #ffffff}A.menu2:visited {color:#FF66cFF;line-height: 25pt;font-size:9pt}</style>"
javastr=javastr+"<table width=\"100%\" border=\"0\">"
<!--#include file="articleconn.asp"-->
<%week1=request("week1")
if week1=2  then
'按点击次数从大到小排序，显示最新10条本周内的新闻
sql="select top 10 * from article where shenghe=1 and dateandtime>=date()-6 order by hits desc"
set rs=conn.execute(sql)%>
<% 
dim i
do while not rs.eof 
%>
javastr=javastr+"<tr><td>"
javastr=javastr+"<font color=\"#0066cc\">○ </font><span style=\"font-size:9pt;line-height: 12pt\"><a href=\"../open.asp?id=<%=rs("newsid")%>&path=<%=rs("path")%>&filename=<%=rs("N_Fname")%>\") ><%=rs("title")%></span></a>"


javastr=javastr+"</td></tr>"
<%i=i+1
	if i=9 then exit do
	rs.movenext
	loop
rs.close   
set rs=nothing   
conn.close   
set conn=nothing %> 
javastr=javastr+"</table>" 
document.write (javastr) 
<%else%>
document.write ("对不起，暂时没有任何内容。")
<%end if%>

