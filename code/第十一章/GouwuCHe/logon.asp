<!--#include file="md5.asp"-->
<%
  '�����δ����Pass�������䶨��ΪFalse����ʾδ��¼
  If Request.Cookies("Order_Info") ="" Then
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
  			"Data Source="&Server.MapPath("STORE.mdb")
	  Conn.Open
	  Sql="select * from Users_Info where UserName='"&UserName &"' and  UserPwd='"&UserPwd &"'"
	   '��ȡ�û�����
 
	  set rs=Conn.Execute(Sql) 
	  If rs.EOF Then  
         '�û������ڣ���ʾ������Ϣ
	    Errmsg = "�û�������"
   	  Else      
           '��¼�ɹ�
          Errmsg = ""
          Response.Cookies("Order_Info")("User")=MD5(UserName)
          Response.Cookies("Order_Info")("ID")=rs.Fields("UserID") 
    	  Response.Cookies("Order_Info").Expires=Date()+1
    	  Session("UserName") = UserName
	      Response.Redirect("index.asp")      
	  End If
    End If
 Else
 	 If IsEmpty(Session("UserName")) Then 
 	 	Response.Cookies("Order_Info")=""
 	Else
  	 	Response.Redirect("index.asp")
  	 End If
 End If
 'δ��¼���ߵ�¼���ɹ�����ʾ��¼����
  If Request.Cookies("Order_Info")="" Then
%>
<HTML>
<HEAD><TITLE>�������û���������</TITLE></HEAD>
<BODY>
<p align="center"><font face="�����п�" size="6" color="#0000FF">��&nbsp; ��&nbsp; ��&nbsp; ģ&nbsp; ��</font></p> 

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