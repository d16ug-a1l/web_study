<!--  
 * ====================================================================
 *
 *                Send.asp ���������߼���֧���ṩ
 *
 *  ��ҳ�����������ҳ���ж�����Ϣ,���ύ֧����������������֧��ƽ̨....
 *
 * 
 * ====================================================================
-->

<!--#include file="MD5.asp"-->

<%
'****************************************	
	v_mid = "20000400"					             ' �̻���,����Ϊ�����̻���20000400���滻Ϊ�Լ����̻��ż���
	v_url = "http://localhost/chinabank/Receive.asp" ' �̻��Զ��巵�ؽ���֧�������ҳ�� Receive.asp Ϊ����ҳ��

													 ' MD5��ԿҪ�������ύҳ��ͬ����Send.asp��� key = "test" ,�޸�""���� test Ϊ������Կ
	key = "test"									 ' �������û������MD5��Կ���½����Ϊ���ṩ�̻���̨����ַ��https://merchant3.chinabank.com.cn/
													 ' ��½��������ĵ�����������ҵ������Ϲ����������Ϲ���Ķ������������С�MD5��Կ���á� 
													 ' ����������һ��16λ���ϵ���Կ����ߣ���Կ���64λ��������16λ�Ѿ��㹻��
'****************************************%>


<%
   if request("v_oid")<>"" then									'�ж��Ƿ��д��ݶ�����
   
		  v_oid = request("v_oid")
	  
	  else

		  curdate = now()										' ����ϵͳʱ�������������ʽ��YYYYMMDD-v_mid-HMMSS
		  ymd = year(curdate)&month(curdate)&day(curdate)		' ������
		  hms = hour(curdate)&minute(curdate)&second(curdate)	' ����ʱ

		  v_oid = ymd&"-"&v_mid&"-"&hms							' �Ƽ������Ź��ɸ�ʽΪ ������-�̻���-Сʱ������

	end if

	v_amount = request("v_amount")		' �������

	v_moneytype = "CNY"					' ����

	text = v_amount&v_moneytype&v_oid&v_mid&v_url&key	' ƴ�ռ��ܴ�

	v_md5info=Ucase(trim(md5(text)))					' ����֧��ƽ̨��MD5ֵֻ�ϴ�д�ַ���������Сд��MD5ֵ��ת��Ϊ��д



%>

<!--������ϢΪ��׼�� HTML ��ʽ + ASP ���� ƴ�ն��ɵ� �������� ֧���ӿڱ�׼��ʾҳ�� -->

<html>

<body onLoad="javascript:document.E_FORM.submit()">
<form action="https://pay3.chinabank.com.cn/PayGate" method="POST" name="E_FORM">

  <!--���¼���Ϊ����֧����Ҫ��Ϣ����Ϣ������ȷ������Ϣ��Ӱ��֧�����У�-->
    
  <input type="hidden" name="v_md5info"    value="<%=v_md5info%>" size="100">
  <input type="hidden" name="v_mid"        value="<%=v_mid%>">
  <input type="hidden" name="v_oid"        value="<%=v_oid%>">
  <input type="hidden" name="v_amount"     value="<%=v_amount%>">
  <input type="hidden" name="v_moneytype"  value="<%=v_moneytype%>">
  <input type="hidden" name="v_url"        value="<%=v_url%>">

    
  <!--���¼���������֧�������޹أ�ֻ��������¼�ͻ���Ϣ�����Բ��ã�ʹ�úͲ�ʹ�ö���Ӱ��֧�� -->

	<input type="hidden"  name="v_rcvname"      value="<%=v_rcvname%>">
	<input type="hidden"  name="v_rcvaddr"      value="<%=v_rcvaddr%>">
	<input type="hidden"  name="v_rcvtel"       value="<%=v_rcvtel%>">
	<input type="hidden"  name="v_rcvpost"      value="<%=v_rcvpost%>">
	<input type="hidden"  name="v_rcvemail"     value="<%=v_rcvemail%>">
	<input type="hidden"  name="v_rcvmobile"    value="<%=v_rcvmobile%>">

	<input type="hidden"  name="v_ordername"    value="<%=v_ordername%>">
	<input type="hidden"  name="v_orderaddr"    value="<%=v_orderaddr%>">
	<input type="hidden"  name="v_ordertel"     value="<%=v_ordertel%>">
	<input type="hidden"  name="v_orderpost"    value="<%=v_orderpost%>">
	<input type="hidden"  name="v_orderemail"   value="<%=v_orderemail%>">
	<input type="hidden"  name="v_ordermobile"  value="<%=v_ordermobile%>">
  
  </form>

</body>
</html>