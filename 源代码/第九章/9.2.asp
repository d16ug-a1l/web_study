<%
'��ȡ�������Ƿ�֧�ָ������Ϣ
str=HaveObj("Scripting.FileSystemObject ")
'�жϷ������Ƿ�֧�ָ����
If str=False Then
	Response.write("��֧��Scripting.FileSystemObject���")
Else
	Response.Wrtie("֧��Scripting.FileSystemObject ���")
End If
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
