  
<%
  on error resume next
  Dim strOS,strHomeDrive,strHomePath,strPath,strWindir,strTemp
  Dim ObjName(13,2)
  ObjName(0,0) = "MSWC.AdRotator"
  ObjName(0,1) = "ϵͳ�Դ�������"
  ObjName(1,0) = "MSWC.BrowserType"
  ObjName(1,1) = "�������Ϣ���" 
  ObjName(2,0) = "MSWC.NextLink"
  ObjName(2,1) = "ϵͳ�Դ��������"
  ObjName(3,0) = "MSWC.Tools"
  ObjName(4,0) = "MSWC.Status"
  ObjName(5,0)= "MSWC.Counters"
  ObjName(5,1) = "ϵͳ�Դ��������"
  ObjName(6,0)= "IISSample.ContentRotator"
  ObjName(6,1) = "ϵͳ�Դ����ݹ�����"
  ObjName(7,0)= "IISSample.PageCounter"
  ObjName(7,1) = "ϵͳ�Դ�ͳ�����"
  ObjName(8,0) = "Microsoft.XMLHTTP"
  ObjName(8,1) = "(Http ���, ���ڲɼ�ϵͳ���õ�)"
  ObjName(9,0) = "WScript.Shell"
  ObjName(9,1) = "(Shell ���, �����漰��ȫ����)"
  ObjName(10,0) = "Scripting.FileSystemObject"
  ObjName(10,1) = "(FSO �ļ�ϵͳ�����ı��ļ���д)"
  ObjName(11,0) = "Adodb.Connection"
  ObjName(11,1) = "(ADO ���ݶ���)"
  ObjName(12,0) = "Adodb.Stream"
  ObjName(12,1) = "(ADO ����������, ����������������ϴ�������)"
  ObjName(13,0) = "JMail.SmtpMail"	
  ObjName(13,1) = "JMail�����ʼ����"
  GetOSInfo
  
  Response.write "����ϵͳΪ��"&strOS&"<BR>"
  Response.write "����������Ϊ��"&strHomeDrive&"<BR>"
  Response.write "�û�Ĭ��·��Ϊ��"&strHomePath&"<BR>"
  Response.write "��������·��Ϊ��"&strPath&"<BR>"
  Response.write "ϵͳĿ¼Ϊ��"&strWindir&"<BR>"
  Response.write "��ʱ�ļ�Ŀ¼Ϊ��"&strTemp&"<BR>"
  
  For i=0 To 13 
  	ObjCheck(ObjName(i,0))
  	If IsObj Then
  		Response.write "ϵͳ֧��"&ObjName(i,0)&"�����"&ObjName(i,1)&"  "&VerObj&"<BR>"
  	Else
  		Response.write "ϵͳ��֧��"&ObjName(i,0)&"�����"&"<BR>"
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
    strOS = "(δ֪)"
  end if
end sub
  %>