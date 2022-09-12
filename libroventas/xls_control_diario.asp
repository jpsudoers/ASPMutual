<!--#include file="../conexion.asp"-->
<%
Server.ScriptTimeout=0
fecha = year(now)&right("0"&month(now()),2)&right("0"&day(now()),2)
fecha = fecha&right("0"&hour(now()),2)&right("0"&minute(now()),2)&right("0"&second(now()),2)
Response.ContentType ="application/vnd.ms-excel"
Response.AddHeader "content-disposition", "inline; filename=Control_Diario_"&fecha&".xls"

	qsHc="SELECT F.FACTURA,CONVERT(VARCHAR(10),F.FECHA_EMISION,105) AS FECHA_EMISION, "
	qsHc=qsHc&"CONVERT(VARCHAR(10),F.FECHA_FACTURA,105) AS FECHA_FACTURA,C.DESCRIPCION,P.DIR_EJEC, "
	qsHc=qsHc&"(CASE WHEN A.ID_PROYECTO IS NULL then 'Minera Escondida' ELSE A.ID_PROYECTO "
	qsHc=qsHc&"END) as 'proyecto',e.RUT,DBO.MayMinTexto(E.R_SOCIAL) AS R_SOCIAL,f.MONTO,a.N_REG_SENCE,(CASE F.ESTADO "
	qsHc=qsHc&"WHEN 1 THEN 'Vigente' WHEN 2 THEN 'Refacturada' WHEN 0 THEN 'Anulada' end) as 'ESTADO' FROM FACTURAS F  "
	qsHc=qsHc&" inner join AUTORIZACION a on a.ID_AUTORIZACION=f.ID_AUTORIZACION "
	qsHc=qsHc&" inner join empresas e on e.ID_EMPRESA=f.ID_EMPRESA "
	qsHc=qsHc&" inner join programa p on p.ID_PROGRAMA=a.ID_PROGRAMA "
	qsHc=qsHc&" inner join CURRICULO c on c.ID_MUTUAL=p.ID_MUTUAL "
	qsHc=qsHc&" WHERE F.FECHA_EMISION>=CONVERT(date, '"&request("i")&"') and  "
	qsHc=qsHc&" F.FECHA_EMISION<=CONVERT(date, '"&request("t")&"') "
	qsHc=qsHc&" order by f.FACTURA asc"

	set rsHc =  conn.execute(qsHc)%>

    <table width="4000" border="1">
    <tr>
      <td width="100" align="center"><b><font size="2">Fecha Generación</font></b></td>
      <td width="100" align="center"><b><font size="2">Fecha Emisión</font></b></td>
      <td width="200" align="center"><b><font size="2">Empresa</font></b></td>
      <td width="400" align="center"><b><font size="2">Sucursal</font></b></td>
      <td width="100" align="center"><b><font size="2">N° Factura</font></b></td>
      <td width="100" align="center"><b><font size="2">Estado</font></b></td>
      <td width="100" align="center"><b><font size="2">N° Reg. Sence</font></b></td>
      <td width="450" align="center"><b><font size="2">Cliente</font></b></td>
      <td width="100" align="center"><b><font size="2">Rut Cliente</font></b></td> 
      <td width="400" align="center"><b><font size="2">Nombre Proyecto</font></b></td>      
      <td width="650" align="center"><b><font size="2">Tipo de Servicio y/o Curso</font></b></td>
      <td width="100" align="center"><b><font size="2">Monto $</font></b></td>
    </tr>
	<%
    while not rsHc.eof
    %>
    <tr>
      <td align="center"><%=rsHc("FECHA_FACTURA")%></td>
	  <td align="center"><%=rsHc("FECHA_EMISION")%></td> 
	  <td align="center">Mutual Capacitación</td>  
	  <td><%=rsHc("DIR_EJEC")%></td>  
	  <td align="center"><%=rsHc("FACTURA")%></td> 
      <td align="center"><%=rsHc("ESTADO")%></td> 
      <td align="center"><%=rsHc("N_REG_SENCE")%></td> 
	  <td><%=rsHc("R_SOCIAL")%></td>  
	  <td align="right"><%=rsHc("RUT")%></td>  
	  <td><%=rsHc("proyecto")%></td> 
	  <td><%=rsHc("DESCRIPCION")%></td>  
	  <td align="right"><%=rsHc("MONTO")%></td>        
    </tr>	
    <%
        rsHc.MoveNext
    wend
    %>
    </table>