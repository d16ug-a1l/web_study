<html>
<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>广告轮流显示的内置组件</title>
</head>
<body>
<%
Dim i
Dim Obj
'创建Content Linking组件
Set Obj=Server.CreateObject("MSWC.NextLink")
'获取当前页面在NextLink.txt文件中的序号
i=Obj.GetListIndex("NextLink.txt")
Dim strContent
'获取当前页面的描述内容
strContent=Obj.GetNthDescription("NextLink.txt",i)
'输出当前页面的描述内容
Response.write("<p align='center'><font face='华文行楷' size='6' color='#0000FF'>"& strContent&"</font></p> ")
Dim strPrev,strPrevURL
Dim strNext,strNextURL
If i=1 Then 
'如果当前页为第一页，则没有前一页链接
'没有此判断，程序会出现错误
	strPrevURL=""
	strPrev="前一页"
Else
'获取前一页的URL
strPrevURL=Obj.GetPreviousURL("NextLink.txt")
'获取前一页的描述内容
strPrev=Obj.GetPreviousDescription("NextLink.txt")
'设置前一页的链接
strPrevURL="<a href='"& strPrevURL&"'>"
End If
If i=Obj.GetListCount("NextLink.txt") Then 
'如果当前页为最好一页，则没有后一页链接
'如果没有此判断，会出现代码错误
	strNext="后一页"
	strNextURL=""
Else
'获取后一页的描述内容
strNext=Obj.GetNextDescription("NextLink.txt")
'获取后一页的链接URL
strNextURL=Obj.GetNextURL("NextLink.txt")
'设置后一页的链接
strNextURL=" <a href='"& strNextURL&"'>"
End If
Response.write("<p align='center'>"&strPrevURL&strPrev&"</a>")
Response.write(strNextURL&strNext&"</a></p> ")
%>
<p>&nbsp;&nbsp;&nbsp;
<span lang="EN-US" style="font-size: 10.5pt; font-family: Times New Roman">AD 
Rotator</span><span style="font-size: 10.5pt; font-family: 宋体">组件是一个广告轮流显示的内置组件，AD Rotato组件能够实现广告图片随机显示到页上，并且可以设定各个广告出现的频率。AD Rotator组件还可以轻松添加或者修改广告的超链接，这样用户就可以通过单击广告访问广告客户的WEB站点。AD Rotator组件简化了轮流显示每幅广告的任务，并且让添加和修改广告的工作变得更轻松。</span></p>
</body>
</html>
