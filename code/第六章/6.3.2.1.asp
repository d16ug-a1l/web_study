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
'����FSO����ʵ������������fso
set fso=Server.CreateObject("Scripting.FileSystemObject")
%>
<table border=1 width="100%">
<tr><td>�̷�</td><td>����</td><td>���</td><td>�ܼƴ�С</td><td>���ÿռ�</td>
<td>�ļ�ϵͳ</td><td>���к�</td><td>�Ƿ����</td><td>·��</td></tr>
<%
'ѭ������ÿһ�����������̣�������������Ϣ����Drives��
For each driver in fso.Drives
	Response.Write "<tr>"
	Response.Write "<td>" & driver.DriveLetter & "</td>"	 ' ����̷�
	Response.write "<td>" & Num2Info(driver.DriveType) & "</td>" '�����������
	'�жϵ�ǰ�����Ƿ�����ã����Ϊ����û�з������ʱ�ǲ������õģ������Ҫ�ж�
	If driver.IsReady Then
		'�������������ꡢ�������������ÿռ䡢�ļ�ϵͳ�����к�
		Response.write "<td>" & driver.VolumeName & "</td>"
		Response.write "<td>" & FormatNumber(driver.TotalSize / 1024, 0)& "</td>" 
		Response.write "<td>" & FormatNumber(driver.Availablespace / 1024, 0) & "</td>" 
		Response.write "<td>" & driver.FileSystem & "</td>" 
		Response.write "<td>" & driver.SerialNumber & "</td>"
	Else
		'���̲������������������Ϣ
		Response.write "<td>��</td>"
		Response.write "<td>�� </td>"
		Response.write "<td>�� </td>"
		Response.write "<td>�� </td>"
		Response.write "<td>�� </td>"
	End If
	Response.write "<td>" & driver.IsReady & "</td>" 
Response.write "<td>" & driver.Path & "</td>"
Response.Write "</tr>"
Next
set fso=nothing
%>
</table>
