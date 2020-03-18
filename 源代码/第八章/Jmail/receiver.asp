<%
User=Trim(Request.Form("Receive")) '创建JMail的POP3对象
PWD=trim(Request.Form("PWD"))
POP3Server =trim(Request.Form("POP3"))
Set pop3 = Server.CreateObject( "JMail.POP3" ) 
'连接POP3服务器。connect的语法格式为connect user,password,POP3Server 
'user：用户名
'password：密码
'POP3Server :POP3服务器地址
pop3.Connect User,PWD, POP3Server 
'显示邮件数目
Response.Write( "You have " & pop3.count & " mails in your mailbox!<br><br>" )
set msg=server.CreateObject("JMail.message")
'获取所有邮件
for i=1 to pop3.count 
Set msg = pop3.Messages.item(i) '读取第I封邮件
separator = ", "
%> 
<html>
<body>
<TABLE>
<tr>
<td>Subject</td>
<td><%
'获取并输出邮件的标题
Response.Write msg.Subject '输出邮件的标题
%></td>
</tr>
<tr>
<td>From</td>
<td><%
'获取并输出邮件的发信人
Response.Write msg.FromName '输出邮件的发信人姓名
%></td>
</tr>
<tr>
<td>Attachments</td>
<td><%
'获取并输出邮件的附件信息
Response.Write getAttachments '输出邮件的附件信息
%></td>
</tr>
<tr>
<td>Body</td>
<td><%
'输出邮件的正文。使用HTMLBody属性才可以正常显示包括文本和HTML格式的内容
Response.Write msg.HTmlBody %></td>
</tr> 
<%
Next
'获取附件信息的函数
Function getAttachments()  
Set Attachments = msg.Attachments '获取邮件的附件
'判断附件的数目
If Attachments.Count>0 Then
separator = ", "
Response.Write "<br>"&msg.size&"<br>" '输出邮件大小
For i = 0 To Attachments.Count - 1
If i = Attachments.Count - 1 Then
separator = ""
End If
Set at = Attachments(i) '获取第I个附件
at.SaveToFile( "c:\" & at.Name )
'输出附件的链接
getAttachments = getAttachments & "<a href=""" & at.Name &""">" &_
at.Name & "(" & at.Size & " bytes)" & "</a>" & separator
Next
End IF
End Function
pop3.Disconnect '断开与POP3邮件服务器的链接
%>
</TABLE>
</body>
</html>
