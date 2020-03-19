<%
dim a,b
'给定密码
a="123456"
'再次输入密码b
b="234567"
if StrComp(a,b)<>0 then
response.write "你输入的密码不正确" 
end if
%> 