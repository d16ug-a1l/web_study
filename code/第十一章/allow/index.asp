<!--#include file="funciton.asp"-->
<body bgcolor="#99CCFF">

<p align="center"><font face="华文行楷" size="6" color="#0000FF">权 限 管 理 模 块</font></p> 
<%

If Not(Session("Pass") = True and  Session("User") <>"" and Session("Id") <>"" and Session("GroupID")<>"" )Then
   Response.Redirect("logon.asp")
End If
UserID=Cint(Session("Id"))
GroupID=Cint(Session("GroupID"))
 'response.write(UserID&"<BR>")
' response.write(GroupID&"<BR>")
%>
<table border="1" width="100%" id="table1" height="500">
			
	<tr valign=top>
		<td width='25%'>
		

		  <table width='100%' >
		  <%
			nCode=Cint(GetResAllow("公告栏",UserID,GroupID))
			'Response.write(nCode)
			If nCode>=1 Then
			%>
		   <tr>
		    <td>
			<table border="0" width="100%" id="table3" height=100>
			<tr>
		    <td><IMG height=20  src="img/TitleSquare.gif" width=20>
			<font face="华文行楷" size="6" color="#0000FF">公告栏</font>
			</td></tr>
				<%
			    OutPutFileContent  "公告栏",UserID,GroupID  			    
				%> 
				 
			</table>
			</td>
			</tr>
  
			<%
			    End If
				nCode=Cint(GetResAllow("信息交流",UserID,GroupID))
				If nCode>=1 Then
			%>
			
			 <tr>
		    <td><IMG height=20  src="img/TitleSquare.gif" width=20>
			<font face="华文行楷" size="6" color="#0000FF">信息交流</font>
			<table border="0" width="100%" id="table3" height=58>
				<%
			    OutPutFileContent  "信息交流",UserID,GroupID  			    
				
				%> 
			</table>

			</td>
			</tr>
			<%
  				End If
				nCode=Cint(GetResAllow("相关链接",UserID,GroupID))
				If nCode>=1 Then
			%>
			 <tr>
		    <td><IMG height=20  src="img/TitleSquare.gif" width=20>
			<font face="华文行楷" size="6" color="#0000FF">相关链接</font>
			<table border="0" width="100%" id="table3" height=58>
				<%
			    OutPutFileContent  "相关链接",UserID,GroupID  			    
				
				%> 
			</table>

			</td>
			</tr>
			<%
			End If
				nCode=Cint(GetResAllow("管理功能",UserID,GroupID))
				'Response.write(nCode )
				If nCode>=1 Then
			%>
			<tr>
		    <td><IMG height=20  src="img/TitleSquare.gif" width=20>
			<a href="admin.asp"><font face="华文行楷" size="6" color="#0000FF">管理功能</font></a>
			
			</td>
			</tr>
			<%
				End If
			%>
		  </table>
		</td>
		<%
				nCode=Cint(GetResAllow("新闻动态",UserID,GroupID))
				If nCode>=1 Then
		%>
		<td width="50%">
		<table width='100%'valign="top" id="table2">
		
		   <tr  >
		    <td valign="top" ><IMG height=20  src="img/TitleSquare.gif" width=20>
			<font face="华文行楷" size="6" color="#0000FF">新闻动态</font>
			<table border="0" width="100%" id="table3" height=100>
				<%
			    OutPutFileContent  "新闻动态",UserID,GroupID 			    
				%> 
			</table>　　
			</td>
			</tr>  
		      <%
		 		End If
		 		%>
		 	</table>
		  </td>
		  <%
				nCode=Cint(GetResAllow("网上调查",UserID,GroupID))
				If nCode>=1 Then
			%>
		　
				
		
		<td width="30%">
		<table width='100%'valign=top>
		   <tr>
		    <td>
		    
			<table border="0" width="100%" id="table3" height=58>
			<tr><td ><IMG height=20  src="img/TitleSquare.gif" width=20>
			<font face="华文行楷" size="6" color="#0000FF">网上调查</font></td></tr>
			<tr><td>对哪一章比较满意？<form method="POST" action="--WEBBOT-SELF--">
				<!--webbot bot="SaveResults" U-File="fpweb:///_private/form_results.csv" S-Format="TEXT/CSV" S-Label-Fields="TRUE" -->
				<p><input type="radio" value="V1" checked name="R1">
				第一章<p><input type="radio" name="R1" value="V2">
				第二章<p><input type="radio" name="R1" value="V3">
				第三章<p><input type="submit" value="提交" name="B1"><input type="reset" value="重置" name="B2">
				</form>
				<p>　</td></tr>
				<%
			    OutPutFileContent  "网上调查",UserID,GroupID  			    
				%> 
			</table>　
			</td>
			</tr>
			<%
				End If
				nCode=Cint(GetResAllow("站内地图",UserID,GroupID))

				If nCode>=1 Then
			%>
			<tr>
		    <td>
			<table border="0" width="100%" id="table3" height=58>
			<tr>
		    <td ><IMG height=20  src="img/TitleSquare.gif" width=20>
			<font face="华文行楷" size="6" color="#0000FF">站内地图</font>
			</td></tr>
				<%
			    OutPutFileContent  "站内地图",UserID,GroupID  			    
				%> 
			</table>　
			</td>
			</tr>
			<%
				End If
				nCode=Cint(GetResAllow("支持服务",UserID,GroupID))
				If nCode>=1 Then
			%>
			<tr>
		    <td><IMG height=20  src="img/TitleSquare.gif" width=20>
			<font face="华文行楷" size="6" color="#0000FF">支持服务</font>　
			<table border="0" width="100%" id="table3" height=58>
				<%
			    OutPutFileContent  "支持服务",UserID,GroupID 			    
				%> 
			</table>　
			</td>
			</tr>
			<%
				End If
			%>
		  </table>
		</td>
	</tr>
</table>
 
