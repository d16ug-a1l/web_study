<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>盘符</title>
</head>

<body>
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
set fso=Server.CreateObject("Scripting.FileSystemObject")
%>

<table border=1 width="100%">
<tr><td>盘符</td><td>类型</td><td>卷标</td><td>总计大小</td><td>可用空间</td>
<td>文件系统</td><td>序列号</td><td>是否可用</td><td>路径</td></tr>
<%  
set fso=Server.CreateObject("Scripting.FileSystemObject")
DriverPath="c"
If fso.DriveExists(DriverPath) Then
	Set drive=fso.GetDrive(DriverPath)
	Response.Write "<tr>"
	Response.Write "<td>" & drive.DriveLetter & "</td>"
	Response.write "<td>" & Num2Info(drive.DriveType) & "</td>"
	If drive.IsReady Then
		Response.write "<td>" & drive.VolumeName & "</td>"
		Response.write "<td>" & FormatNumber(drive.TotalSize / 1024, 0)& "</td>" 
		Response.write "<td>" & FormatNumber(drive.Availablespace / 1024, 0) & "</td>" 
		Response.write "<td>" & drive.FileSystem & "</td>" 
		Response.write "<td>" & drive.SerialNumber & "</td>"
	Else
		Response.write "<td>无</td>"
		Response.write "<td>无 </td>"
		Response.write "<td>无 </td>"
		Response.write "<td>无 </td>"
		Response.write "<td>无 </td>"
	End If
	Response.write "<td>" & drive.IsReady & "</td>" 
	Response.write "<td>" & drive.Path & "</td>" 		
	Response.Write "</tr>"
End If
set fso=nothing
%>
</table>

</body>

</html>
