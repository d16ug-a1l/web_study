<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�½���ҳ 1</title>
</head>

<body>
<%
SMTP=Trim(Request.Form("SMTP"))
If SMTP="" Then
	Response.write("�����ʼ�����������Ϊ�գ�")
	Response.End
End If
PWD=Request.Form("PWD")
strTo=Request.Form("Receive")
If strTo="" Then
	Response.write("�������ʼ���ַ����Ϊ�գ�")
	Response.End
End If
CC=Request.Form("CC")
BCC=Request.Form("BCC")
If CC<>"" Then
JMail.AddRecipientCC CC
End If
If BCC<>"" Then
JMail.AddRecipientBCC BCC
End If
From=Request.Form("Sender")
Title=Request.Form("Title")
Msg=Request.Form("Content")
Attatchment=Request.Form("Attatchment")
 
HTML=Request.Form("C1")
Set JMail=Server.CreateObject("JMail.Message")
JMail.Silent  =  True
JMail.AddHeader "Originating-IP", Request.ServerVariables("REMOTE_ADDR")  
JMail.From=From
mails=split(strTo,",")
For each mail in mails
	JMail.AddRecipient mail 
Next
JMail.Subject=Title
If HTML="ON" Then
	JMail.HTMLBody =Msg
	JMail.AppendHTML "man��JMail�ʼ�����ϵͳ"
Else
JMail.Body  =Msg
JMail.AppendText "man��JMail�ʼ�����ϵͳ"
End If
If Attatchment<>"" Then
	JMail.ContentType  =  "Multipart/mixed" 
Else
	JMail.ContentType  =  "text/html"  
End If
JMail.logging=true
JMail.Charset  =  "gb2312"  
JMail.Priority = 3

If Trim(Attatchment)<>"" Then
JMail.AddAttachment Attatchment, true
End If

Dim n
n=Instr(From,"@")
If n>0 Then
Sender=Mid(From,1,n-1)
JMail.MailServerUserName  = From
JMail.MailServerPassWord  =  PWD
JMail.MailServerUserName =SMTP
Response.write sender&"<BR>" 
Response.Write  Sender&":"&PWD&"@"&SMTP&"<BR>"
err=JMail.Send("asp_man:"&PWD&"@"&SMTP) 
If Trim(Jmail.Errorcode)<>""Then 
 Response.Write "ERR CODE is "&Jmail.Errorcode&"<BR>"
End If
If Trim(Jmail.errormessage)<>""Then 
 Response.Write "ERR Message is "&Jmail.errormessage&"<BR>"
End If
If Trim(Jmail.errorsource)<>""Then 
 Response.Write "Err Source is "&Jmail.errorsource&"<BR>"
End IF
if  err  then    
     SendMail=  err.description  
     err.clear  
else  
     SendMail="���ͳɹ�"  
end  if  

 
Response.Write SendMail  
Else
Response.Write "�����˵ĵ�ַ����ȷ"
End if 
JMail.Close  
set  JMail=  nothing  
%>
</body>

</html>
