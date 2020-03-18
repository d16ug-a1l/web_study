<%
  Book = Request("Book")   
  Relax = Request("Relax")  
  Cost= Request("Cost")  
  Num = Request("Quantity")   
  subTotal = Cost* Num  
%>
<HTML>
  <BODY >
 
    <%
    '新增记录
     Set Conn=Server.CreateObject("ADODB.Connection")
      'Response.Write(Server.MapPath("User.mdb")&"<BR>")
	  Conn.ConnectionString="Provider=Microsoft.Jet.OLEDB.4.0;"&_
  			"Data Source="&Server.MapPath("STORE.mdb")
	  Conn.Open
	  Sql="Select ID from shop_list where ShopName='"&trim(Session("UserName"))&"'"
	  set rs=Conn.Execute(Sql)
	  if rs.EOF=False Then
	  	Response.write("该商品已经购买过，请删除原有商品后重新购买！")
	  Else
	  	Sql="Insert into [shop_list]([User],[ShopName],[Relax],[Cost],[TotalCost],[Time]) "&_
		 " values('"&Session("UserName")&"','"&Book&"','"&Relax &"',"&Cost&","&subTotal &",'"&Now()&"')"    
     	set rs=Conn.Execute(Sql)    
      End If
    '关闭数据库连接并释放对象
 
    %>
 
    <CENTER>　<P>选取的商品已放入购物车！<P>
    <P><A href="Catalog.asp">返回产品类型</A></P></CENTER>                                                                               
  </BODY>
</HTML>