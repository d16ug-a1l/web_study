<HTML>
  <BODY BGCOLOR="LightYellow">
    <%  
      ID = Request("ID")                                                                         
      Set Conn=Server.CreateObject("ADODB.Connection")
      'Response.Write(Server.MapPath("User.mdb")&"<BR>")
	  Conn.ConnectionString="Provider=Microsoft.Jet.OLEDB.4.0;"&_
  			"Data Source="&Server.MapPath("STORE.mdb")
	  Conn.Open
      Sql= "Delete * FROM shop_list WHERE ID=" & ID    
       set rs=Conn.Execute(Sql)           
    %>
    <!-显示标题图片、成功信息及返回购物车的超链接 -->
    <CENTER>　<P>选取的商品已删除！<P>
    <P><A HREF="Car.asp">返回购物车</A></P></CENTER>
  </BODY>                                                                                 
</HTML>