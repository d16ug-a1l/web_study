<HTML>
  <BODY BACKGROUND="bg1.jpg">
    <H3>注意事项</H3>
    <OL TYPE="1">
      <LI>订阅方法一：本网站使用网银在线支付。</LI>
      <LI>订阅方法二：请利用邮局汇款单，填妥姓名、户名、商品、数目，直接至邮局邮购付款。帐号：*****户名：*****</LI>
    </OL><HR>
    <%
     Set Conn=Server.CreateObject("ADODB.Connection")
      'Response.Write(Server.MapPath("User.mdb")&"<BR>")
	  Conn.ConnectionString="Provider=Microsoft.Jet.OLEDB.4.0;"&_
  			"Data Source="&Server.MapPath("STORE.mdb")
	  Conn.Open
      Sql= "SELECT * FROM shop_list WHERE user='" & Session("UserName") & "'"    
      set rs=Conn.Execute(Sql) 
      Dim Name
      If rs.EOF Then 
        Response.Write "<CENTER><IMG SRC='fig1.jpg'><P>购物车内没任何商品！</P>" & _
          "<P><A HREF='Catalog.asp'>产品类型</A></P></CENTER>" 
      Else 
      	  Total = 0  
 
	      Do While Not rs.EOF
 
	        Name=Name&rs("ShopName")&","
	        Total = Total + rs("TotalCost")  
	        rs.MoveNext
      Loop
      
      End If
      Response.write("<form method='post' action='chinabank/send.asp?v_amount="&Total&"'>")
   %>
   
    <TABLE BORDER="1"  RULES="Cols" ALIGN="Center" CELLPADDING="5">
      <TR HEIGHT="25"> <TD ALIGN="Center" BGCOLOR="#CCCC00">订单信息</TD></TR>
      <TR HEIGHT="25"><TD>商品名称：<U><%= Name %></U></TD></TR>
      <TR HEIGHT="25"><TD>货款：<U><%= Total  %></U></TD></TR> 
      </TABLE>
 	 <p align="center"> <input type="submit" value="提交" name="B1"></p>
 	 <p align="center"><a href="car.asp">返回购物车</a>  </p>

</form>
  
  </BODY>
</HTML>