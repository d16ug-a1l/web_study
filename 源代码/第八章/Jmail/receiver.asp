<%
User=Trim(Request.Form("Receive")) '����JMail��POP3����
PWD=trim(Request.Form("PWD"))
POP3Server =trim(Request.Form("POP3"))
Set pop3 = Server.CreateObject( "JMail.POP3" ) 
'����POP3��������connect���﷨��ʽΪconnect user,password,POP3Server 
'user���û���
'password������
'POP3Server :POP3��������ַ
pop3.Connect User,PWD, POP3Server 
'��ʾ�ʼ���Ŀ
Response.Write( "You have " & pop3.count & " mails in your mailbox!<br><br>" )
set msg=server.CreateObject("JMail.message")
'��ȡ�����ʼ�
for i=1 to pop3.count 
Set msg = pop3.Messages.item(i) '��ȡ��I���ʼ�
separator = ", "
%> 
<html>
<body>
<TABLE>
<tr>
<td>Subject</td>
<td><%
'��ȡ������ʼ��ı���
Response.Write msg.Subject '����ʼ��ı���
%></td>
</tr>
<tr>
<td>From</td>
<td><%
'��ȡ������ʼ��ķ�����
Response.Write msg.FromName '����ʼ��ķ���������
%></td>
</tr>
<tr>
<td>Attachments</td>
<td><%
'��ȡ������ʼ��ĸ�����Ϣ
Response.Write getAttachments '����ʼ��ĸ�����Ϣ
%></td>
</tr>
<tr>
<td>Body</td>
<td><%
'����ʼ������ġ�ʹ��HTMLBody���Բſ���������ʾ�����ı���HTML��ʽ������
Response.Write msg.HTmlBody %></td>
</tr> 
<%
Next
'��ȡ������Ϣ�ĺ���
Function getAttachments()  
Set Attachments = msg.Attachments '��ȡ�ʼ��ĸ���
'�жϸ�������Ŀ
If Attachments.Count>0 Then
separator = ", "
Response.Write "<br>"&msg.size&"<br>" '����ʼ���С
For i = 0 To Attachments.Count - 1
If i = Attachments.Count - 1 Then
separator = ""
End If
Set at = Attachments(i) '��ȡ��I������
at.SaveToFile( "c:\" & at.Name )
'�������������
getAttachments = getAttachments & "<a href=""" & at.Name &""">" &_
at.Name & "(" & at.Size & " bytes)" & "</a>" & separator
Next
End IF
End Function
pop3.Disconnect '�Ͽ���POP3�ʼ�������������
%>
</TABLE>
</body>
</html>
