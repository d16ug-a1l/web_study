<form method="POST" action="search.asp">
	<p align="center"><font face="�����п�" size="6" color="#0000FF">�� �� ģ ��</font></p>
	<p align="center"><font face="�����п�" color="#0000FF">�ֶ�</font><select size="1" name="Field">
	<%
	  If IsEmpty( Session("prev_search")) Then
	  	Session("prev_search")=""
	  End If
	  Set Conn=Server.CreateObject("ADODB.Connection")
      'Response.Write(Server . MapPath("User.mdb")&"<BR>")
	  Conn.ConnectionString="Provider=Microsoft.Jet.OLEDB.4.0;"&_
  			"Data Source="&Server.MapPath("book.mdb")
	  Conn.Open
	  Sql="select * from Book_Info"
	  set rs=Conn.Execute(Sql) 
	  If rs.EOF=false Then  
         For i=0 To rs.Fields.Count-1
         	Response.write("<option value='"&rs.Fields(i).Name&","&rs.Fields(i).type&"'>"&rs.Fields(i).Name&"</option>")
         Next   	  	 
   	  End If
	%>
	</select><font face="�����п�" color="#0000FF">����</font>
	<input type="text" name="Content" size="16">
	<select size="1" name="Type">
	<option selected value="MH">ģ��</option>
	<option value="JQ">��ȷ</option>
	<option value="DY">����</option>
	<option value="XY">С��</option>
	<option value="NOT">������</option>
	</select><input type="submit" value="����" name="B1"></p>
	<p align="center"><input type="radio" value="Reset" checked name="R1">��������<input type="radio" name="R1" value="Result">�ڽ���в�ѯ</p>
</form>
<%
  Content=trim(Request.Form("Content"))
  If Content<>"" Then
	  Fields=trim(Request.Form("Field"))
	  str=split(Fields,",")
	  Dim Field,VarType
	  If IsArray(str) Then
		  Field=str(0)
		  VarType=Cint(str(1))
	  End If
	  LinkType=Request.Form("Type")
	  
	  If LinkType="MH" Then
	  	link="like"
	  ElseIf LinkType="DY" Then
	  	link=">"
	  ElseIf LinkType="XY" Then
	  	link="<"
	  ElseIf LinkType="NOT" Then
	  	link="<>"
	  Else
	  	link="="
	  End If
	  'Response.write(link)
	  Sql="Select * from book_Info "
	  Dim SearchSql,bType
	  bType=0
	  If VarType=3 Then
	  	If IsNumeric(Content) Then
	  		SearchSql =Field&link&CInt(Content)
	  	Else
	  		bType=1
	  	End If
	  ElseIf VarType=202 Then
	  	SearchSql =Field&" "&link&" '%"&Content&"%'"
	  ElseIf VarType=7 Then
	    If IsDate(Content) Then
	  		SearchSql =" datediff('d','"&Content&"',data)"&link&"0"
	  	Else 
	  		bType=1
	  	End If
	  End If
 
	  If bType=0 Then
	 	  Result=Request.Form("R1")
		  If Result="Result" and Session("prev_search")<>"" Then
		  	Sql=Sql&Session("prev_search")&" and "&SearchSql 
		  Else  
		  	SearchSql=" where "& SearchSql
		  	Sql=Sql&SearchSql
		  End If
 
		  Session("prev_search")=SearchSql
 
		  Sql=Sql&" order by ID"
		  Set Conn=Server.CreateObject("ADODB.Connection")
 
		  Conn.ConnectionString="Provider=Microsoft.Jet.OLEDB.4.0;"&_
		 			"Data Source="&Server.MapPath("book.mdb")
		  Conn.Open
		  'Response.write(Sql)
		 set rs=Conn.Execute(Sql) 
		 Response.write("<table align=center> ")
		 Response.write("<tr><td ><font face='�����п�' color='#0000FF'>��ѯ���</font></td></tr>")
		 Dim nCount
		 nCount=0
		 Do while rs.EOF=false 
		   Response.write(" <tr><td>"&rs("Name")&"</td></tr>") 	  	 
		   rs.movenext
		   nCount=nCount+1
		 loop
		 If nCount=0 Then
		 	Response.write("<tr><td ><font face='�����п�' color='#FF0000'>���޲�ѯ���</font></td></tr>")
		 End If
		 Response.write("</table> ")
	End If
 End If
%>