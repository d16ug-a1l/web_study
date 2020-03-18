<!--  
 * ====================================================================
 *
 *                Send.asp 由网银在线技术支持提供
 *
 *  本页面接收来自上页所有订单信息,并提交支付订单到网线在线支付平台....
 *
 * 
 * ====================================================================
-->

<!--#include file="MD5.asp"-->

<%
'****************************************	
	v_mid = "20000400"					             ' 商户号,这里为测试商户号20000400，替换为自己的商户号即可
	v_url = "http://localhost/chinabank/Receive.asp" ' 商户自定义返回接收支付结果的页面 Receive.asp 为接收页面

													 ' MD5密钥要跟订单提交页相同，如Send.asp里的 key = "test" ,修改""号内 test 为您的密钥
	key = "test"									 ' 如果您还没有设置MD5密钥请登陆我们为您提供商户后台，地址：https://merchant3.chinabank.com.cn/
													 ' 登陆后在上面的导航栏里可能找到“资料管理”，在资料管理的二级导航栏里有“MD5密钥设置” 
													 ' 建议您设置一个16位以上的密钥或更高，密钥最多64位，但设置16位已经足够了
'****************************************%>


<%
   if request("v_oid")<>"" then									'判断是否有传递订单号
   
		  v_oid = request("v_oid")
	  
	  else

		  curdate = now()										' 根据系统时间产生订单，格式：YYYYMMDD-v_mid-HMMSS
		  ymd = year(curdate)&month(curdate)&day(curdate)		' 年月日
		  hms = hour(curdate)&minute(curdate)&second(curdate)	' 分秒时

		  v_oid = ymd&"-"&v_mid&"-"&hms							' 推荐订单号构成格式为 年月日-商户号-小时分钟秒

	end if

	v_amount = request("v_amount")		' 订单金额

	v_moneytype = "CNY"					' 币种

	text = v_amount&v_moneytype&v_oid&v_mid&v_url&key	' 拼凑加密串

	v_md5info=Ucase(trim(md5(text)))					' 网银支付平台对MD5值只认大写字符串，所以小写的MD5值得转换为大写



%>

<!--以下信息为标准的 HTML 格式 + ASP 语言 拼凑而成的 网银在线 支付接口标准演示页面 -->

<html>

<body onLoad="javascript:document.E_FORM.submit()">
<form action="https://pay3.chinabank.com.cn/PayGate" method="POST" name="E_FORM">

  <!--以下几项为网上支付重要信息，信息必须正确无误，信息会影响支付进行！-->
    
  <input type="hidden" name="v_md5info"    value="<%=v_md5info%>" size="100">
  <input type="hidden" name="v_mid"        value="<%=v_mid%>">
  <input type="hidden" name="v_oid"        value="<%=v_oid%>">
  <input type="hidden" name="v_amount"     value="<%=v_amount%>">
  <input type="hidden" name="v_moneytype"  value="<%=v_moneytype%>">
  <input type="hidden" name="v_url"        value="<%=v_url%>">

    
  <!--以下几项与网上支付货款无关，只是用来记录客户信息，可以不用，使用和不使用都不影响支付 -->

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