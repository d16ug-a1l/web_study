<!--#include file="funciton.asp"-->
<body bgcolor="#99CCFF">

<p align="center"><font face="�����п�" size="6" color="#0000FF">Ȩ �� �� �� ģ ��</font></p> 
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
			nCode=Cint(GetResAllow("������",UserID,GroupID))
			'Response.write(nCode)
			If nCode>=1 Then
			%>
		   <tr>
		    <td>
			<table border="0" width="100%" id="table3" height=100>
			<tr>
		    <td><IMG height=20  src="img/TitleSquare.gif" width=20>
			<font face="�����п�" size="6" color="#0000FF">������</font>
			</td></tr>
				<%
			    OutPutFileContent  "������",UserID,GroupID  			    
				%> 
				 
			</table>
			</td>
			</tr>
  
			<%
			    End If
				nCode=Cint(GetResAllow("��Ϣ����",UserID,GroupID))
				If nCode>=1 Then
			%>
			
			 <tr>
		    <td><IMG height=20  src="img/TitleSquare.gif" width=20>
			<font face="�����п�" size="6" color="#0000FF">��Ϣ����</font>
			<table border="0" width="100%" id="table3" height=58>
				<%
			    OutPutFileContent  "��Ϣ����",UserID,GroupID  			    
				
				%> 
			</table>

			</td>
			</tr>
			<%
  				End If
				nCode=Cint(GetResAllow("�������",UserID,GroupID))
				If nCode>=1 Then
			%>
			 <tr>
		    <td><IMG height=20  src="img/TitleSquare.gif" width=20>
			<font face="�����п�" size="6" color="#0000FF">�������</font>
			<table border="0" width="100%" id="table3" height=58>
				<%
			    OutPutFileContent  "�������",UserID,GroupID  			    
				
				%> 
			</table>

			</td>
			</tr>
			<%
			End If
				nCode=Cint(GetResAllow("������",UserID,GroupID))
				'Response.write(nCode )
				If nCode>=1 Then
			%>
			<tr>
		    <td><IMG height=20  src="img/TitleSquare.gif" width=20>
			<a href="admin.asp"><font face="�����п�" size="6" color="#0000FF">������</font></a>
			
			</td>
			</tr>
			<%
				End If
			%>
		  </table>
		</td>
		<%
				nCode=Cint(GetResAllow("���Ŷ�̬",UserID,GroupID))
				If nCode>=1 Then
		%>
		<td width="50%">
		<table width='100%'valign="top" id="table2">
		
		   <tr  >
		    <td valign="top" ><IMG height=20  src="img/TitleSquare.gif" width=20>
			<font face="�����п�" size="6" color="#0000FF">���Ŷ�̬</font>
			<table border="0" width="100%" id="table3" height=100>
				<%
			    OutPutFileContent  "���Ŷ�̬",UserID,GroupID 			    
				%> 
			</table>����
			</td>
			</tr>  
		      <%
		 		End If
		 		%>
		 	</table>
		  </td>
		  <%
				nCode=Cint(GetResAllow("���ϵ���",UserID,GroupID))
				If nCode>=1 Then
			%>
		��
				
		
		<td width="30%">
		<table width='100%'valign=top>
		   <tr>
		    <td>
		    
			<table border="0" width="100%" id="table3" height=58>
			<tr><td ><IMG height=20  src="img/TitleSquare.gif" width=20>
			<font face="�����п�" size="6" color="#0000FF">���ϵ���</font></td></tr>
			<tr><td>����һ�±Ƚ����⣿<form method="POST" action="--WEBBOT-SELF--">
				<!--webbot bot="SaveResults" U-File="fpweb:///_private/form_results.csv" S-Format="TEXT/CSV" S-Label-Fields="TRUE" -->
				<p><input type="radio" value="V1" checked name="R1">
				��һ��<p><input type="radio" name="R1" value="V2">
				�ڶ���<p><input type="radio" name="R1" value="V3">
				������<p><input type="submit" value="�ύ" name="B1"><input type="reset" value="����" name="B2">
				</form>
				<p>��</td></tr>
				<%
			    OutPutFileContent  "���ϵ���",UserID,GroupID  			    
				%> 
			</table>��
			</td>
			</tr>
			<%
				End If
				nCode=Cint(GetResAllow("վ�ڵ�ͼ",UserID,GroupID))

				If nCode>=1 Then
			%>
			<tr>
		    <td>
			<table border="0" width="100%" id="table3" height=58>
			<tr>
		    <td ><IMG height=20  src="img/TitleSquare.gif" width=20>
			<font face="�����п�" size="6" color="#0000FF">վ�ڵ�ͼ</font>
			</td></tr>
				<%
			    OutPutFileContent  "վ�ڵ�ͼ",UserID,GroupID  			    
				%> 
			</table>��
			</td>
			</tr>
			<%
				End If
				nCode=Cint(GetResAllow("֧�ַ���",UserID,GroupID))
				If nCode>=1 Then
			%>
			<tr>
		    <td><IMG height=20  src="img/TitleSquare.gif" width=20>
			<font face="�����п�" size="6" color="#0000FF">֧�ַ���</font>��
			<table border="0" width="100%" id="table3" height=58>
				<%
			    OutPutFileContent  "֧�ַ���",UserID,GroupID 			    
				%> 
			</table>��
			</td>
			</tr>
			<%
				End If
			%>
		  </table>
		</td>
	</tr>
</table>
 
