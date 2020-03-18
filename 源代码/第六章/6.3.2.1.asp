<%
'使用函数Num2Info()把驱动器编码转换成驱动器的说明信息
Function Num2Info(Driver)
Select Case Driver
Case 0: Num2Info="设备无法识别"
Case 1: Num2Info="软盘驱动器"
Case 2: Num2Info="硬盘驱动器"
Case 3: Num2Info="网络硬盘驱动器"
Case 4: Num2Info="光盘驱动器"
Case 5: Num2Info="RAM虚拟磁盘"
End Select
End Function
'创建FSO对象实例并赋给变量fso
set fso=Server.CreateObject("Scripting.FileSystemObject")
%>
<table border=1 width="100%">
<tr><td>盘符</td><td>类型</td><td>卷标</td><td>总计大小</td><td>可用空间</td>
<td>文件系统</td><td>序列号</td><td>是否可用</td><td>路径</td></tr>
<%
'循环处理每一个驱动器磁盘，所有驱动器信息存在Drives中
For each driver in fso.Drives
	Response.Write "<tr>"
	Response.Write "<td>" & driver.DriveLetter & "</td>"	 ' 输出盘符
	Response.write "<td>" & Num2Info(driver.DriveType) & "</td>" '输出磁盘类型
	'判断当前磁盘是否可以用，如果为光驱没有放入光盘时是不可以用的，因此需要判断
	If driver.IsReady Then
		'以上依次输出卷标、磁盘容量、可用空间、文件系统和序列号
		Response.write "<td>" & driver.VolumeName & "</td>"
		Response.write "<td>" & FormatNumber(driver.TotalSize / 1024, 0)& "</td>" 
		Response.write "<td>" & FormatNumber(driver.Availablespace / 1024, 0) & "</td>" 
		Response.write "<td>" & driver.FileSystem & "</td>" 
		Response.write "<td>" & driver.SerialNumber & "</td>"
	Else
		'磁盘不可以用则输出以下信息
		Response.write "<td>无</td>"
		Response.write "<td>无 </td>"
		Response.write "<td>无 </td>"
		Response.write "<td>无 </td>"
		Response.write "<td>无 </td>"
	End If
	Response.write "<td>" & driver.IsReady & "</td>" 
Response.write "<td>" & driver.Path & "</td>"
Response.Write "</tr>"
Next
set fso=nothing
%>
</table>
