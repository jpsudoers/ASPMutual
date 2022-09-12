<!--#include file="../conexion.asp"-->
<%
fecha = year(now)&right("0"&month(now()),2)&right("0"&day(now()),2)
fecha = fecha&right("0"&hour(now()),2)&right("0"&minute(now()),2)&right("0"&second(now()),2)
Response.ContentType ="application/vnd.ms-excel"
Response.AddHeader "content-disposition", "inline; filename=emp_"&fecha&".xls"
	

sql = "SELECT distinct e.* FROM FACTURAS F "
sql = sql&" inner join EMPRESAS E on E.ID_EMPRESA=F.ID_EMPRESA "
sql = sql&" where /*f.id_factura in (40352,40353,40355,40359,40360,40361,40363)*/"
sql = sql&" e.rut like '%78311400%' or "
sql = sql&" e.rut like '%83710400%' or "
sql = sql&" e.rut like '%95616000%'"



	set rs =  conn.execute(sql)%>

    <table width="4000" border="1">
    <tr>
          <td align="left"><b><font size="2">rut</font></b></td>
          <td><b><font size="2">R_SOCIAL</font></b></td>
          <td><b><font size="2">GIRO</font></b></td>          
          <td><b><font size="2">DIRECCION</font></b></td>
          <td><b><font size="2">COMUNA</font></b></td>    
          <td><b><font size="2">CIUDAD</font></b></td>  
          <td align="right"><b><font size="2">FONO</font></b></td>
         
    </tr>	    
	<%
    while not rs.eof
    %>
    <tr>
          <td align="left"><b><font size="2"><%=rs("rut")%></font></b></td>
          <td><b><font size="2"><%=rs("R_SOCIAL")%></font></b></td>
          <td><b><font size="2"><%=rs("GIRO")%></font></b></td>          
          <td><b><font size="2"><%=rs("DIRECCION")%></font></b></td>
          <td><b><font size="2"><%=rs("COMUNA")%></font></b></td>    
          <td><b><font size="2"><%=rs("CIUDAD")%></font></b></td>  
          <td align="right"><b><font size="2"><%=rs("FONO")%></font></b></td>

    </tr>	
    <%
        rs.MoveNext
    wend
    %>
    </table>
