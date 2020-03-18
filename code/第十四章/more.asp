<!--#include file="head.asp" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title><%=title%>_<%=request("sss")%></title>
<link href="css.css" rel="stylesheet" type="text/css">
</head>
<body leftmargin="0" topmargin="0">
<table width="770" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td width="20">&nbsp;</td>
    <td width="529" valign="top">您的位置： <a href="index.asp">首页</a> - 
      <%sss=request("sss")%> <span class="unnamed1"><%=sss%></span></td>
    <td width="20">&nbsp;</td>
    <td rowspan="4" align="right" valign="top"> <table width="161" border="0" cellpadding="3" cellspacing="1" bgcolor="#6687BA">
        <tr> 
          <td width="161" height="20" background="images/bg11.gif"> <div align="center">站 
              内 搜 索</div></td>
        </tr>
        <tr> 
          <td bgcolor="#F2F4F9"> <form name="searchtitle" onsubmit="if(searchtitle.ttt.value.length<1){alert('搜索关键字不能为空！：）');return(false)}else{return(true)}" method="POST" action="search.asp" target="_blank">
              <div align="center"> 
                <input name="ttt" type="text" class="unnamed5" style="FONT-SIZE: 9pt" onfocusin='vbscript:searchtitle.ttt.value=""' value="请输入关键字" size="16">
                <br>
                <select class="unnamed5" name="sss" size="1" style="FONT-SIZE: 9pt">
                  <option selected>按新闻标题搜索</option>
                  <option>按新闻内容搜索</option>
                </select>
                <br>
                <input type="submit" name="Submit" value="搜 索" class="unnamed5" style="FONT-SIZE: 9pt">
                <input type="reset" name="Submit2" value="取 消" class="unnamed5" style="FONT-SIZE: 9pt">
              </div>
            </form></td>
        </tr>
        <tr> 
          <td height="20" background="images/bg11.gif"> <div align="center">推 
              荐 信 息</div></td>
        </tr>
        <tr> 
          <td height="10" bgcolor="#F2F4F9"><script language="javascript" src="tjnews1.asp?tjnews=1"></script></td>
        </tr>
        <tr> 
          <td height="20" background="images/bg11.gif"> <div align="center">本 
              周 热 点</div></td>
        </tr>
        <tr> 
          <td height="9" bgcolor="#F2F4F9"><script language="javascript" src="week1.asp?week1=2"></script></td>
        </tr>
        <tr> 
          <td height="10" bgcolor="#F2F4F9">&nbsp;</td>
        </tr>
      </table></td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
    <td rowspan="2" valign="top"> <!--#include file="articleconn.asp"--> <!--#include file="chkstr.inc"--> <% 
mmm=request("mmm")
if mmm="" then mmm=0'设置初始页数
ttt=request("ttt")'得到栏目编号
'查询数据库得到此栏目编号所对应的所有文章信息
set rs=server.createobject("adodb.recordset")  
sql ="select * from article where (typeid like '%"&checkStr(ttt)&"%') order by dateandtime Desc" 
rs.open sql,conn,1,1%> <center>
      </center>
      <% if rs.eof and rs.bof then  
response.write "<p align='center'>【<a href='javascript:window.close()'>关闭窗口</a>】"
response.end
end if 
i=0 %> <br> <table width="98%" border="0" cellpadding="4" cellspacing="0" bordercolor="#000000" bordercolorlight="#000000" bordercolordark="#FFFFFF" bgcolor="#DBDBDB" class="unnamed2">
        <td width="453"><form method=Post action="search.asp">
              <%    
'下面是分页显示      
  if mmm<>0 then         
  	for iisf=1 to mmm *5       
  		if rs.eof then exit for         
  		rs.movenext         
  	next         
  end if         
  do while not rs.eof          
  %>
            </form>
        <tr bgcolor="#FFFFFF"> 
          <td width="453" align="left">&nbsp;<a href="<%=rs("path")%>/<%=rs("N_Fname")%>">○<b> 
            </b><%=rs("title")%></a> </td>
          <td width="160" bgcolor="#FFFFFF" align="center"><%=rs("dateandtime")%></td>
        </tr>
        <% i=i+1                                       
   rs.movenext                                       
   if i=20 then exit do                                       
   loop               
%>
        <td width="453"><form method=Post action="search.asp">
          </form>
      </table>
      <p align="center"> <span class="unnamed1"> 
        <!--上页-->
        <%if mmm<>0 then%>
        <%="<a href=more.asp?mmm=" & mmm-1 & "&sss=" & sss & "&ttt=" & ttt & ">上一页</a>"%> 
        <%end if%>
        <!--下页-->
        <%if not rs.eof then%>
        <%="<a href=more.asp?mmm=" & mmm+1 & "&sss=" & sss & "&ttt=" & ttt & ">下一页</a>"%> 
        <%end if%>
        </span></p>
      <% 
rs.close              
set rs=nothing              
conn.close              
set conn=nothing %> <br> </td>
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
    <td valign="top">&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
</table>
</body>
</html>
<!--#include file="topy.asp" -->