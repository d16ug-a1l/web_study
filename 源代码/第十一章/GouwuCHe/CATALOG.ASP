<HTML>
  <BODY >
    <%   
      Set Conn=Server.CreateObject("ADODB.Connection")
      'Response.Write(Server.MapPath("User.mdb")&"<BR>")
	  Conn.ConnectionString="Provider=Microsoft.Jet.OLEDB.4.0;"&_
  			"Data Source="&Server.MapPath("STORE.mdb")
	  Conn.Open
	
      Sql= "SELECT * FROM Type"    
      set rs=Conn.Execute(Sql) 
    %>

    <TABLE BORDER="0" ALIGN="Center" WIDTH="90%"> 
      <TR BGCOLOR='#ACACFF'  HEIGHT="30" ALIGN="Center">
        <%
          '读取数据表的字段名称以作为表格的标题
          For I = 0 To rs.Fields.Count - 1
            Response.Write "<TD>" & rs.Fields(I).Name & "</TD>"
          Next
          Response.Write "<TD>数量</TD>"
          Response.Write "<TD>订购</TD>"
        %>
      </TR>
    <%
    '读取各个字段的数据并显示在表格内
    Do While Not rs.EOF
      Data = "<TR HEIGHT='30' BGCOLOR='#EAEAFF'>"
      For I = 0 To rs.Fields.Count - 1
        Data = Data & "<TD>" & rs.Fields(I).Value & "</TD>"
      Next
      Response.Write Data
      Response.Write "<TD><FORM METHOD='POST' TARGET= 'Bottom' ACTION='AddToCar.asp?Book=" & _
        rs("Name") & "&Relax=" & rs("Relax") & "&Cost=" & rs("Cost") & _
        "'><INPUT TYPE='TEXT' NAME='Quantity' SIZE='5'></TD>"
      Response.Write "<TD><INPUT TYPE='SUBMIT' VALUE='放入购物车'></TD></FORM></TR>"
      rs.MoveNext
    Loop
    
    '关闭数据库连接并释放对象
    rs.Close         
    Set rs = Nothing 
    Conn.Close         
    Set Conn = Nothing 
    %>
  </TABLE>
  </BODY>                                                                                 
</HTML>