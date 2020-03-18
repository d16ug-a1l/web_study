<%
'获取服务器是否支持该组件信息
str=HaveObj("Scripting.FileSystemObject ")
'判断服务器是否支持该组件
If str=False Then
	Response.write("不支持Scripting.FileSystemObject组件")
Else
	Response.Wrtie("支持Scripting.FileSystemObject 组件")
End If
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
