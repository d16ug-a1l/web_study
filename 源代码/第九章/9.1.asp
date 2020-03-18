<%
'strFileContent为用户提交的文件内容
Function IncludeKeyWord(strFileContent)
	'HaveKey表示文件内容是否包含关键字
	'为False表示没有特定关键字，为True为包含
HaveKey=false
'设定的关键字
	strKeyWords="server|createobject|execute|encode|eval|request|activexobject|language=" 
	strKeyWord=split(strKeyWords,"|")
	'转换成小写，防止用户把关键字写成大写蒙混过关
strFileContent=LCase(strFileContent)
'检测提交内容是否含有特定关键字
	For i=0 to ubound(strKeyWord)
		'返回特定关键字在提交内容中的位置
		n=Instr(strFileContent,strKeyWord(i))
		if n>0 Then
			'如果包含“server”，则该关键字后必须为“.”
			If Instr(strKeyWord(i),"server")>0 Then
				m=n+6
				'去除“server”后面的空格
				Do while trim(Mid(strFileContent,m,1))="" and m<=len(strFileContent)-1
					m=m+1
				loop
				'检测下一个字符是否为“.”
				If Mid(strFileContent,m,1)="." Then
					'包含“server.”终止循环，设置HaveKey
HaveKey=true
					Exit For
				End If
			End If
			'检测是否包含“.createobject”和“. encode”
			If Instr(strKeyWord(i),"createobject")>0 or Instr(strKeyWord(i),"encode")>0 Then
				m=n-1
				'去除关键字前面的空格
				Do while trim(Mid(strFileContent,m,1))="" and m>1 
					m=m-1
				loop
				'判断前面的字符是“.”
				If Mid(strFileContent,m,1)="." Then
					HaveKey=true
					Exit For
				End If
			End If
		Exit For
		End If
	Next
	'设置该函数的返回值
	IncludeKeyWord=HaveKey
End Function
%>
