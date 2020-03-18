<%
response.write "服务器地址:"&Request.ServerVariables("SERVER_NAME")&"<BR>"
response.write "服务器IP:"&Request.ServerVariables("LOCAL_ADDR")&"<BR>"
response.write "IIS版本:"&Request.ServerVariables("SERVER_SOFTWARE")&"<BR>"
Set WshShell = server.CreateObject("WScript.Shell")
Set WshSysEnv = WshShell.Environment("SYSTEM")
okOS = cstr(WshSysEnv("OS"))
response.write "服务器CPU信息:"&okOS &"<BR>"
'定义组件
dim ZJ(10)
ZJ(10) = "MSWC.AdRotator"
ZJ(1) = "MSWC.BrowserType"
ZJ(2) = "MSWC.NextLink"
ZJ(3) = "MSWC.Tools"
ZJ(4) = "MSWC.Status"
ZJ(5) = "MSWC.Counters"
ZJ(6) = "IISSample.ContentRotator"
ZJ(7) = "IISSample.PageCounter"
ZJ(8) = "MSWC.PermissionChecker"
ZJ(9) = "Microsoft.XMLHTTP"
for i=1 to 10 
'获取服务器是否支持该组件信息
str=HaveObj(ZJ(i))
'判断服务器是否支持该组件
If str=False Then
	Response.write("不支持"&ZJ(i)&"组件"&"<BR>")
Else
	Response.Write("支持"&ZJ(i)&"组件"&"<BR>")
End If
next
'判断服务器是否支持指定的组件
'这是一种办法。在Server一章也介绍了另外一种办法，读者可以参考8.2.2。
Function HaveObj(strObj)
  '启动错误处理程序
  on error resume next
  Dim Have		'保存创建组件是否成功信息。True为创建组件成功
  Have=false
  Dim str
  str =""
'创建该组件
  set Obj=server.CreateObject (strObj)
'判断是否出现错误
  If -2147221005 <> Err then 
     Have = True
	'获取该组件信息
     str = Obj.version
     if str ="" or isnull(str) then str = Obj.about
  end if
  set TestObj=nothing
  If Have Then
'获取该组件的信息
	HaveObj=str
  Else
'返回不支持信息
	HaveObj=false
  End If
End Function
%>
