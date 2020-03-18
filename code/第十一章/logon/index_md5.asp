<%
<!--#include files="md5.asp"-->
UserPwd="adsf"
user=MD5(UserPwd)
Response.write(user)
%>
