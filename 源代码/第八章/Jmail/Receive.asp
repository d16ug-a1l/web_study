<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>新建网页 1</title>
</head>

<body>
<%
SMTP=Trim(Request.Form("SMTP"))
If SMTP="" Then
	Response.write("发送邮件服务器不能为空！")
	Response.End
End If
PWD=Request.Form("PWD")
strTo=Request.Form("Receive")
If strTo="" Then
	Response.write("接收人邮件地址不能为空！")
	Response.End
End If
From=Request.Form("Sender")
Title=Request.Form("Title")
Msg=Request.Form("Content")
Add=Request.Form("Add")
response.write(Add)
HTML=Request.Form("C1")
Set JMail=Server.CreateObject("JMail.Message")
'Response.Write IsObjInstalled("JMail.Message")
Response.Write("saddfafds")
JMail.From=From
mails=split(strTo,",")
For each mail in mails
	JMail.AddRecipient strTo
Next
JMail.Subject=Title
If HTML="ON" Then
	JMail.HTMLBody =Msg
	JMail.AppendHTML "<H2>man的JMail邮件测试系统</H2>"
Else
JMail.Body  =Msg
JMail.AppendText "<H2>man的JMail邮件测试系统</H2>"
End If
JMail.ContentType  =  "text/html"  
JMail.Charset  =  "gb2312"  
JMail.Priority = 3
'【参数设置是(True)否(False)为Inline方式】
If Trim(Add)<>"" Then
JMail.AddAttachment Server.MapPath(Add), True 
End If
'JMail.MailServerUserName  =  "man_zl"   
'JMail.MailServerPassWord  =  ""
err=  JMail.Send("man_zl:manzl0228@sina.com.cn") 
Response.Write jmail.log

if  err  then    
     SendMail=  err.description  
     err.clear  
else  
     SendMail="发送成功"  
end  if  
JMail.Close  
set  JMail=  nothing  
 
Response.Write SendMail  
Function IsObjInstalled(strClassString)
On Error Resume Next
IsObjInstalled = False
Err = 0
Dim xTestObj
Set xTestObj = Server.CreateObject(strClassString)
If 0 = Err Then IsObjInstalled = True
Set xTestObj = Nothing
Err = 0
Response.End()
End Function 
%>
</body>

</html>
