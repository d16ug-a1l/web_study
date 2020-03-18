<!--#include file="top.asp"-->
<%keyword=trim(request("keyword"))
KEYWORD=replace(keyword,"'","")
if keyword="" then
         	response.write"<SCRIPT language=JavaScript>alert('错误，没有正确填写关键字！');"
			response.write"javascript:history.go(-1)</SCRIPT>"
			response.end
end if
sql="select * from ly where  title like '%"&keyword&"%' order by tt desc"
set rs= Server.CreateObject("ADODB.RecordSet") 
rs.Open sql, Conn, 1, 2%>
<table width="98%" border="0" cellspacing="0" cellpadding="0" align="center" style="BORDER-LEFT: #000000 1px solid; BORDER-RIGHT: #000000 1px solid" bgcolor="#64B3D9" >
  <tr>
    <td> <br>
      <br>&nbsp;&nbsp;以下为搜索“<%=keyword%>”的结果
      <table border="0" width="98%" cellspacing="1" cellpadding="0"  bgcolor="#205E7B" align="center" >
        <tr> 
          <td colspan="2" align="center" height="25"><b><font color="#FFFFFF">帖子主题</font></b></td>
          <td align="center" width="84"><b><font color="#FFFFFF">作&nbsp;者</font></b></td>
          <td align="center" width="69"><b><font color="#FFFFFF">回复/阅读</font></b></td>
          <td align="center" width="196" ><b><font color="#FFFFFF">最后回复时间/回复人</font></b></td>
        </tr>
<%


page=Request("page")
  if page=0 then
     page=1
  end if
RecordCount = 0 
  do while not rs.Eof
    RecordCount = RecordCount +1
    rs.MoveNext 
  loop
  if not RecordCount=0 then
     rs.MoveFirst
  end if
  pageCount=RecordCount/25
  pageCount=int(pageCount)
  if (RecordCount mod 25)>0 then
     PageCount=PageCount +1
  end if 
  if pagecount=0 then  pagecount=1
if page>pagecount then page=pagecount
do while pos<(page-1)*25                                 
pos=pos+1                                 
rs.moveNext                                 
loop
x=0
do while x<25 and not rs.eof%>
        <tr bgcolor="#FFFFFF"> 
          <td align="center" bgcolor="#EFEFEF" width="30"> 
            <%if rs("jh")=1 then
      response.write"<img src=""images/isbest.gif"" alt=""精华的帖子""> "
      else
      response.write"<img src=""images/folder.gif"" alt=""普通的帖子"">"
      end if%>
          </td>
          <td height="25" width="*" onMouseOver="this.bgColor = '#f0f0f0'" onMouseOut="this.bgColor = '#ffffff'" > 
            <p><a href="show.asp?id=<%=rs("id")%>"><%=rs("title")%></a> 
          </td>
          <td align="center" bgcolor="#EFEFEF" width="84"> 
            <% if rs("bl")=1 then %>
            <%=rs("user_name")%> 
            <%else%>
            <%=rs("user_name")%> 
            <%end if%>
          </td>
          <td align="center" width="69"><%=rs("hf")%>/<%=rs("hits")%></td>
          <td bgcolor="#EFEFEF" width="195">&nbsp;&nbsp;<%=rs("tt")%>|<%=rs("hfren")%></td>
        </tr>
        <%x=x+1  
  rs.movenext    
loop
rs.close
set rs=nothing
%>
      </table>
      <table width="98%" border="0" align="center" cellspacing="0" cellpadding="3" bgcolor="#205E7B">
        <form method=post action=search1.asp>
          <tr> 
            <td width="52%" class=td1>页次： 
              <%if page="" then
              response.write"1" 
              else
               response.write  page
                end if%>
              /<%=pageCount%> 页 主题数：<%=RecordCount%></td>
            <td width="48%" align="right" class=td1> 
              <%if page=1 then
              response.write"首页" 
              else
               response.write"<a href='search.asp?keyword="&keyword&"' class=a1>首页</a> "
                end if%>
              <%if page=1 then %>
              上一页
              <%else%>
              <a href="search1.asp?page=<%=page-1%>&amp;keyword=<%=keyword%>" class=a1>上一页</a> 
              <%end if%>
              <% y=1%>
              <%do while y<pagecount+1%>
              <a href="search1.asp?page=<%=y%>&amp;keyword=<%=keyword%>" class=a1><%=y%></a> 
              <%                                                 
y=y+1                                    
loop%>
              <%if cint(page)=pagecount then %>
             下一页
              <%else%>
              <a href="search1.asp?page=<%=page+1%>&amp;keyword=<%=keyword%>" class="a1">下一页</a> 
              <%end if%>
              转到: 
<INPUT onMouseOver="this.style.backgroundColor = '#E5F0FF'" style="BORDER-RIGHT: #b4b4b4 1px double; BORDER-TOP: #b4b4b4 1px double; BORDER-LEFT: #b4b4b4 1px double; COLOR: #8888aa; BORDER-BOTTOM: #b4b4b4 1px double; BACKGROUND-COLOR: #ffffff" onMouseOut="this.style.backgroundColor = ''" maxLength=3 size=3 name=page value="<%=page%>">
              页 
              <input onMouseOver="this.style.backgroundColor='#FFC864'" style="BORDER-RIGHT: 1px solid; BORDER-TOP: 1px solid; BORDER-LEFT: 1px solid; COLOR: #000000; BORDER-BOTTOM: 1px solid; HEIGHT: 18px; BACKGROUND-COLOR: #f3f3f3" onMouseOut="this.style.backgroundColor='#f3f3f3'" type=submit value=GO name=submit>
<input type=hidden name="keyword" value="<%=keyword%>">
            </td>
          </tr>
        </form>
      </table>
      <br>
      <table cellspacing=1 cellpadding=3 width="98%" bgcolor="#205E7B" align="center">
        <tr> 
    <td ><font color="#FFFFFF">　-=&gt; <b>BBS图例</b></font></td>
  </tr>
  <tr> 
    <td colspan=2 bgcolor="#FFFFFF"> 
      <table cellspacing=4 cellpadding=0 width="92%" align=center border=0 bgcolor="#FFFFFF">
        <tr> 
          <td><img src="images/folder.gif"> 普通帖子</td>
          <td><img src="images/hotfolder.gif"> 热门帖子</td>
          <td><img src="images/lockfolder.gif"> 锁定的帖子</td>
          <td><img src="images/istop.gif"> 固顶帖子 </td>
          <td><img src="images/isbest.gif"> 精华帖子 </td>
        </tr>
      </table>
    </td>
  </tr>
</table>
      <br>
    </td>
  </tr>
</table>
<table width="98%" border="0" cellspacing="0" cellpadding="0" align="center" bgcolor="#000000">
  <tr>
    <td align="center" height="30"class=td1><%=footer%>
        </td>
  </tr>
</table>
</body>
</html>
