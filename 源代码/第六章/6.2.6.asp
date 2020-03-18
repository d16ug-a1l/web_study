  
<%
  on error resume next
  Dim strOS,strHomeDrive,strHomePath,strPath,strWindir,strTemp
  Dim ObjName(13,2)
  ObjName(0,0) = "MSWC.AdRotator"
  ObjName(0,1) = "系统自带广告组件"
  ObjName(1,0) = "MSWC.BrowserType"
  ObjName(1,1) = "浏览器信息组件" 
  ObjName(2,0) = "MSWC.NextLink"
  ObjName(2,1) = "系统自带链接组件"
  ObjName(3,0) = "MSWC.Tools"
  ObjName(4,0) = "MSWC.Status"
  ObjName(5,0)= "MSWC.Counters"
  ObjName(5,1) = "系统自带计数组件"
  ObjName(6,0)= "IISSample.ContentRotator"
  ObjName(6,1) = "系统自带内容广告组件"
  ObjName(7,0)= "IISSample.PageCounter"
  ObjName(7,1) = "系统自带统计组件"
  ObjName(8,0) = "Microsoft.XMLHTTP"
  ObjName(8,1) = "(Http 组件, 常在采集系统中用到)"
  ObjName(9,0) = "WScript.Shell"
  ObjName(9,1) = "(Shell 组件, 可能涉及安全问题)"
  ObjName(10,0) = "Scripting.FileSystemObject"
  ObjName(10,1) = "(FSO 文件系统管理、文本文件读写)"
  ObjName(11,0) = "Adodb.Connection"
  ObjName(11,1) = "(ADO 数据对象)"
  ObjName(12,0) = "Adodb.Stream"
  ObjName(12,1) = "(ADO 数据流对象, 常见被用在无组件上传程序中)"
  ObjName(13,0) = "JMail.SmtpMail"	
  ObjName(13,1) = "JMail发送邮件组件"
  GetOSInfo
  
  Response.write "操作系统为："&strOS&"<BR>"
  Response.write "本地驱动器为："&strHomeDrive&"<BR>"
  Response.write "用户默认路径为："&strHomePath&"<BR>"
  Response.write "环境变量路径为："&strPath&"<BR>"
  Response.write "系统目录为："&strWindir&"<BR>"
  Response.write "临时文件目录为："&strTemp&"<BR>"
  
  For i=0 To 13 
  	ObjCheck(ObjName(i,0))
  	If IsObj Then
  		Response.write "系统支持"&ObjName(i,0)&"组件。"&ObjName(i,1)&"  "&VerObj&"<BR>"
  	Else
  		Response.write "系统不支持"&ObjName(i,0)&"组件。"&"<BR>"
  	End If
  Next

sub ObjCheck(strObj)
 on error resume next
  IsObj=false
  VerObj=""
  set Obj=server.CreateObject(strObj)
  If IsObject(Obj) then
    IsObj = True
    VerObj =Obj.version
    if VerObj="" or isnull(VerObj) then VerObj=Obj.about
  end if
  set Obj=nothing
End sub	

sub GetOSInfo()
 on error resume next
  Set WshShell = Server.CreateObject("WScript.Shell")
  Set WshEnv = WshShell.Environment("SYSTEM")
  strOS = cstr(WshEnv("OS"))
  strHomeDrive=cstr(WshEnv("HOMEDRIVE"))
  strHomePath=cstr(WshEnv("HOMEPATH"))
  strPath=cstr(WshEnv("PATH"))
  strWindir=cstr(WshEnv("SYSTEMROOT"))
  strTemp=cstr(WshEnv("TEMP"))
  if strOS & "" = "" then
    strOS = "(未知)"
  end if
end sub
  %>