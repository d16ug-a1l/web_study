<HTML>
  <BODY >

    <%  
      Set Conn=Server.CreateObject("ADODB.Connection")
      'Response.Write(Server.MapPath("User.mdb")&"<BR>")
	  Conn.ConnectionString="Provider=Microsoft.Jet.OLEDB.4.0;"&_
  			"Data Source="&Server.MapPath("STORE.mdb")
	  Conn.Open
      Sql= "SELECT * FROM shop_list WHERE user='" & Session("UserName") & "'"    
       set rs=Conn.Execute(Sql) 
     
      If rs.EOF Then 
        Response.Write "<CENTER><P>购物车内没任何商品！</P>" & _
          "<P><A HREF='Catalog.asp'>产品类型</A></P></CENTER>" 
      Else 
    %>  
    <TABLE BORDER="0" ALIGN="Center"> 
      <TR BGCOLOR="#ACACFF" HEIGHT="30" ALIGN="Center">
        <%

          For I = 0 To rs.Fields.Count - 1
            Response.Write "<TD>" & rs.Fields(I).Name & "</TD>"
          Next
          Response.Write "<TD>&nbsp;</TD>"
        %>
      </TR>
    <%
      Total = 0  
      Do While Not rs.EOF
        Data = "<TR HEIGHT='30' BGCOLOR='#EAEAFF'>"
        For I = 0 To rs.Fields.Count - 1
          Data = Data & "<TD>" & rs.Fields(I).Value & "</TD>"
        Next
        Response.Write Data
        Response.Write "<TD><A HREF='Delete.asp?ID=" & rs("ID") & "'>删除</A></TD></TR>"
        Total = Total + rs("TotalCost")  
        rs.MoveNext
      Loop
    

      rs.Close         
      Set rs = Nothing 
      Conn.Close         
      Set Conn = Nothing 
    %>
      <CAPTION ALIGN="Right">总金额：<%= Total %></CAPTION> 
    </TABLE>
    <% End If %>
  </BODY>                                                                                 
</HTML>