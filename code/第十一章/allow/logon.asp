<%
  '�����δ����Pass�������䶨��ΪFalse����ʾδ��¼
  If IsEmpty (Session("Pass")) Then
    Session("Pass") = False
  End If
  '��һ��ִ�иô���
  If Session("Pass")=False  Then
	'��ȡ�ӱ����ݹ������û���������
    UserName = Request.Form("UserName")
	UserPwd = Request.Form("UserPwd")
	'�û���Ϊ�գ���ʾ������Ϣ
    If UserName = "" Then
      Errmsg = "�������û���������!"
    Else    
	  '�������ݿ�
      'Server�����CreateObject��������Connection����
      Set Conn=Server.CreateObject("ADODB.Connection")
      'Response.Write(Server.MapPath("User.mdb")&"<BR>")
	  Conn.ConnectionString="Provider=Microsoft.Jet.OLEDB.4.0;"&_
  			"Data Source="&Server.MapPath("user.mdb")
	  Conn.Open
	  Sql="select * from User_Info where user='"&UserName &"' and  pwd='"&UserPwd &"'"
	   '��ȡ�û�����
	  set rs=Conn.Execute(Sql) 
	  If rs.EOF Then  
         '�û������ڣ���ʾ������Ϣ
	    Errmsg = "�û�������"
   	  Else      
           '��¼�ɹ�
           Errmsg = ""
	      Session("Pass") = True 
	      Session("User") = trim(rs.Fields("user"))
	      Session("Id") = rs.Fields("ID")
	      Session("GroupID") = rs.Fields("GroupID")
	      Response.Redirect("index.asp")
     		
	  End If
    End If
  End If
 'δ��¼���ߵ�¼���ɹ�����ʾ��¼����
  If Session("Pass")=False Then
%>
<HTML>
<HEAD><TITLE>�������û���������</TITLE></HEAD>
<BODY bgcolor="#66CCFF">
<p align="center"><font face="�����п�" size="6" color="#0000FF">�� ¼ ģ ��</font></p> 

<p align="center"><font color="#800000">��<%=Errmsg%></font></p>
<form method="POST" action="logon.asp" name="Form" >
  <p align="center">�û�����&nbsp; <input type="text" name="UserName" size="20"></p>
  <p align="center">��&nbsp; �룺&nbsp; <input type="password" name="UserPwd" size="20"></p>
  <p align="center"><input type="submit" value="�ύ" name="B1"><input type="reset" value="ȫ����д" name="B2"></p>
</form>
<p align="center">��</p>
</BODY>
</HTML>
<%
    Response.End
  End If
%>