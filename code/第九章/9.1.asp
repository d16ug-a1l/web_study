<%
'strFileContentΪ�û��ύ���ļ�����
Function IncludeKeyWord(strFileContent)
	'HaveKey��ʾ�ļ������Ƿ�����ؼ���
	'ΪFalse��ʾû���ض��ؼ��֣�ΪTrueΪ����
HaveKey=false
'�趨�Ĺؼ���
	strKeyWords="server|createobject|execute|encode|eval|request|activexobject|language=" 
	strKeyWord=split(strKeyWords,"|")
	'ת����Сд����ֹ�û��ѹؼ���д�ɴ�д�ɻ����
strFileContent=LCase(strFileContent)
'����ύ�����Ƿ����ض��ؼ���
	For i=0 to ubound(strKeyWord)
		'�����ض��ؼ������ύ�����е�λ��
		n=Instr(strFileContent,strKeyWord(i))
		if n>0 Then
			'���������server������ùؼ��ֺ����Ϊ��.��
			If Instr(strKeyWord(i),"server")>0 Then
				m=n+6
				'ȥ����server������Ŀո�
				Do while trim(Mid(strFileContent,m,1))="" and m<=len(strFileContent)-1
					m=m+1
				loop
				'�����һ���ַ��Ƿ�Ϊ��.��
				If Mid(strFileContent,m,1)="." Then
					'������server.����ֹѭ��������HaveKey
HaveKey=true
					Exit For
				End If
			End If
			'����Ƿ������.createobject���͡�. encode��
			If Instr(strKeyWord(i),"createobject")>0 or Instr(strKeyWord(i),"encode")>0 Then
				m=n-1
				'ȥ���ؼ���ǰ��Ŀո�
				Do while trim(Mid(strFileContent,m,1))="" and m>1 
					m=m-1
				loop
				'�ж�ǰ����ַ��ǡ�.��
				If Mid(strFileContent,m,1)="." Then
					HaveKey=true
					Exit For
				End If
			End If
		Exit For
		End If
	Next
	'���øú����ķ���ֵ
	IncludeKeyWord=HaveKey
End Function
%>
