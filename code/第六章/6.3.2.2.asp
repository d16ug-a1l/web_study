<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�̷�</title>
</head>

<body>
<%
'ʹ�ú���Num2Info()������������ת������������˵����Ϣ
Function Num2Info(Driver)
Select Case Driver
Case 0: Num2Info="�豸�޷�ʶ��"
Case 1: Num2Info="����������"
Case 2: Num2Info="Ӳ��������"
Case 3: Num2Info="����Ӳ��������"
Case 4: Num2Info="����������"
Case 5: Num2Info="RAM�������"
End Select
End Function
set fso=Server.CreateObject("Scripting.FileSystemObject")
%>

<table border=1 width="100%">
<tr><td>�̷�</td><td>����</td><td>���</td><td>�ܼƴ�С</td><td>���ÿռ�</td>
<td>�ļ�ϵͳ</td><td>���к�</td><td>�Ƿ����</td><td>·��</td></tr>
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
		Response.write "<td>��</td>"
		Response.write "<td>�� </td>"
		Response.write "<td>�� </td>"
		Response.write "<td>�� </td>"
		Response.write "<td>�� </td>"
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
