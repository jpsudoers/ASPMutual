<!--#include file="../conexion.asp"-->
<%
fecha = year(now)&right("0"&month(now()),2)&right("0"&day(now()),2)
fecha = fecha&right("0"&hour(now()),2)&right("0"&minute(now()),2)&right("0"&second(now()),2)
Response.ContentType ="application/vnd.ms-excel"
Response.AddHeader "content-disposition", "inline; filename=Bloqueos_"&fecha&".xls"
		
	qsHc="SELECT [estado] ,[portal_norte] ,[portal_sur] ,[rutaux] ,[nomaux] "
	qsHc=qsHc&",[movtipdocref],[movnumdocref],[fecha_emision] ,[fecha_vcto]  "
	qsHc=qsHc&",[debe],[haber] ,[saldo],[fecha_vencicmiento],[dias_vencimiento]  "
    qsHc=qsHc&" FROM dbarauco.dbo.V_SALDOS_PENDIENTES; "
	
	set rsHc =  conn.execute(qsHc)%>

    <table border="1">
    <tr>
      <td align="center"><b><font size="2">Estado</font></b></td>
      <td align="center"><b><font size="2">Portal Norte</font></b></td>      
      <td align="center"><b><font size="2">Portal Sur</font></b></td>   
      <td align="center"><b><font size="2">RUT</font></b></td>
      <td align="center"><b><font size="2">Nombre</font></b></td>
      <td align="center"><b><font size="2">Tipo Doc</font></b></td>      
        <td align="center"><b><font size="2">Num Doc</font></b></td>      
        <td align="center"><b><font size="2">Fecha Emi</font></b></td>      
        <td align="center"><b><font size="2">Fecha Vcto</font></b></td>      
        <td align="center"><b><font size="2">Debe</font></b></td>      
        <td align="center"><b><font size="2">Haber</font></b></td>      
        <td align="center"><b><font size="2">Saldo</font></b></td>      
        <td align="center"><b><font size="2">Fecha V</font></b></td>      
        <td align="center"><b><font size="2">Dias Mora</font></b></td>      
    </tr>
	<%
    while not rsHc.eof
    %>
    <tr>
      <td><font size="2"><%=rsHc("estado")%></font></td>
      <td><font size="2"><%=rsHc("portal_norte")%></font></td>     
      <td><font size="2"><%=rsHc("portal_sur")%></font></td>                  
      <td><font size="2"><%=rsHc("rutaux")%></font></td> 
      <td><font size="2"><%=rsHc("nomaux")%></font></td>
      <td><font size="2"><%=rsHc("movtipdocref")%></font></td>
        <td><font size="2"><%=rsHc("movnumdocref")%></font></td>
        <td><font size="2"><%=rsHc("fecha_emision")%></font></td>
        <td><font size="2"><%=rsHc("fecha_vcto")%></font></td>
        <td><font size="2"><%=rsHc("debe")%></font></td>
        <td><font size="2"><%=rsHc("haber")%></font></td>
        <td><font size="2"><%=rsHc("saldo")%></font></td>
        <td><font size="2"><%=rsHc("fecha_vencicmiento")%></font></td>
        <td><font size="2"><%=rsHc("dias_vencimiento")%></font></td>
    </tr>	
    <%
        rsHc.MoveNext
    wend
    %>
    </table>
