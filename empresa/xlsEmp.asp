<!--#include file="../conexion.asp"-->
<%
fecha = year(now)&right("0"&month(now()),2)&right("0"&day(now()),2)
fecha = fecha&right("0"&hour(now()),2)&right("0"&minute(now()),2)&right("0"&second(now()),2)
Response.ContentType ="application/vnd.ms-excel"
Response.AddHeader "content-disposition", "inline; filename=LST_EMP_"&fecha&".xls"

	qs="select e.rut,UPPER (e.R_SOCIAL) as R_SOCIAL,UPPER (e.GIRO) as GIRO,"&_
       " UPPER (e.DIRECCION) as DIRECCION,UPPER (e.COMUNA) as COMUNA,UPPER (e.CIUDAD) as CIUDAD,"&_
       " e.FONO,e.FAX,UPPER (e.NOMBRES) as EC_NOMBRES,UPPER (e.EMAIL) as EC_EMAIL,"&_
       " e.FONO_CONTACTO AS EC_FONO,UPPER (e.CARGO) as EC_CARGO,"&_
       " UPPER (e.NOMBRE_CONTA) as FN_NOMBRES,UPPER (e.EMAIL_CONTA) as FN_EMAIL,"&_
       " e.FONO_CONTABILIDAD AS FN_FONO,UPPER (e.CARGO_CONTA) as FN_CARGO,CONVERT(VARCHAR(10),e.FECHA_SOLICITUD, 105) AS F_SOLICITUD,"&_
       " (CASE WHEN e.ID_OTIC<>0 THEN (select UPPER (o.R_SOCIAL) from EMPRESAS o "&_
       " where o.ID_EMPRESA=e.ID_OTIC) else 'SIN OTIC' END) as 'OTIC',"&_
       " (CASE WHEN e.MUTUAL<>0 THEN (select UPPER (o.R_SOCIAL) from EMPRESAS o "&_
       " where o.ID_EMPRESA=e.MUTUAL) else 'SIN MUTUAL' END) as 'MUTUAL' "&_
       " from EMPRESAS e "&_
       " where e.tipo=1 and e.estado=1 AND e.PREINSCRITA=1"&_
       " order by e.R_SOCIAL asc"

	set rs =  conn.execute(qs)%>
    <table border="1">
    <tr>
      <td width="80" align="center"><b><font size="2">RUT</font></b></td>
      <td width="400" align="center"><b><font size="2">RAZON SOCIAL</font></b></td>
      <td width="300" align="center"><b><font size="2">GIRO</font></b></td>
      <td width="400" align="center"><b><font size="2">DIRECCION</font></b></td>  
      <td width="200" align="center"><b><font size="2">COMUNA</font></b></td>
      <td width="200" align="center"><b><font size="2">CIUDAD</font></b></td>    
      <td width="200" align="center"><b><font size="2">FONO</font></b></td>    
      <td width="200" align="center"><b><font size="2">FAX</font></b></td>         
      <td width="200" align="center"><b><font size="2">MUTUAL</font></b></td>  
      <td width="650" align="center"><b><font size="2">OTIC</font></b></td>     
      <td width="300" align="center"><b><font size="2">NOMBRE COORDINADOR DE CURSOS</font></b></td>    
      <td width="250" align="center"><b><font size="2">EMAIL COORDINADOR DE CURSOS</font></b></td>  
      <td width="200" align="center"><b><font size="2">FONO COORDINADOR DE CURSOS</font></b></td>     
      <td width="300" align="center"><b><font size="2">CARGO COORDINADOR DE CURSOS</font></b></td>   
      <td width="300" align="center"><b><font size="2">NOMBRE CONTACTO FINANZAS</font></b></td>    
      <td width="250" align="center"><b><font size="2">EMAIL CONTACTO FINANZAS</font></b></td>  
      <td width="200" align="center"><b><font size="2">FONO CONTACTO FINANZAS</font></b></td>     
      <td width="300" align="center"><b><font size="2">CARGO CONTACTO FINANZAS</font></b></td> 
      <td width="140" align="center"><b><font size="2">FECHA DE SOLICITUD</font></b></td>      
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
          <td align="right"><b><font size="2"><%=rs("FAX")%></font></b></td>
          <td><b><font size="2"><%=rs("MUTUAL")%></font></b></td>    
          <td><b><font size="2"><%=rs("OTIC")%></font></b></td>     
          <td><b><font size="2"><%=rs("EC_NOMBRES")%></font></b></td>    
          <td><b><font size="2"><%=rs("EC_EMAIL")%></font></b></td>    
          <td align="right"><b><font size="2"><%=rs("EC_FONO")%></font></b></td>    
          <td><b><font size="2"><%=rs("EC_CARGO")%></font></b></td>  
          <td><b><font size="2"><%=rs("FN_NOMBRES")%></font></b></td>    
          <td><b><font size="2"><%=rs("FN_EMAIL")%></font></b></td>    
          <td align="right"><b><font size="2"><%=rs("FN_FONO")%></font></b></td>    
          <td><b><font size="2"><%=rs("FN_CARGO")%></font></b></td> 
          <td align="center"><b><font size="2"><%=rs("F_SOLICITUD")%></font></b></td>                                              
        </tr>
		<%
	rs.MoveNext
wend
%>
</table>