<%
response.write "��������ַ:"&Request.ServerVariables("SERVER_NAME")&"<BR>"
response.write "������IP:"&Request.ServerVariables("LOCAL_ADDR")&"<BR>"
response.write "IIS�汾:"&Request.ServerVariables("SERVER_SOFTWARE")&"<BR>"
Set WshShell = server.CreateObject("WScript.Shell")
Set WshSysEnv = WshShell.Environment("SYSTEM")
okOS = cstr(WshSysEnv("OS"))
response.write "������CPU��Ϣ:"&okOS &"<BR>"
'�������
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
'��ȡ�������Ƿ�֧�ָ������Ϣ
str=HaveObj(ZJ(i))
'�жϷ������Ƿ�֧�ָ����
If str=False Then
	Response.write("��֧��"&ZJ(i)&"���"&"<BR>")
Else
	Response.Write("֧��"&ZJ(i)&"���"&"<BR>")
End If
next
'�жϷ������Ƿ�֧��ָ�������
'����һ�ְ취����Serverһ��Ҳ����������һ�ְ취�����߿��Բο�8.2.2��
Function HaveObj(strObj)
  '�������������
  on error resume next
  Dim Have		'���洴������Ƿ�ɹ���Ϣ��TrueΪ��������ɹ�
  Have=false
  Dim str
  str =""
'���������
  set Obj=server.CreateObject (strObj)
'�ж��Ƿ���ִ���
  If -2147221005 <> Err then 
     Have = True
	'��ȡ�������Ϣ
     str = Obj.version
     if str ="" or isnull(str) then str = Obj.about
  end if
  set TestObj=nothing
  If Have Then
'��ȡ���������Ϣ
	HaveObj=str
  Else
'���ز�֧����Ϣ
	HaveObj=false
  End If
End Function
%>
