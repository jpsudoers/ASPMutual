<!--#include file="../conexion.asp"-->
<%
Server.ScriptTimeout=0
fecha = year(now)&right("0"&month(now()),2)&right("0"&day(now()),2)
fecha = fecha&right("0"&hour(now()),2)&right("0"&minute(now()),2)&right("0"&second(now()),2)
Response.ContentType ="application/vnd.ms-excel"
Response.AddHeader "content-disposition", "inline; filename=softland_"&fecha&".xls"

	qsHc="SELECT 'F' as tipo_doc,'E' as sub_tipo,(Case F.ESTADO "
	qsHc=qsHc&" When 1 then 'V' else 'N' End) as estado_doc,'22' as bodega,"
	qsHc=qsHc&" F.FACTURA,CONVERT(VARCHAR(10),F.FECHA_EMISION,105) AS FECHA_FACTURA,'01' as conc_doc,"
	qsHc=qsHc&" e.RUT,C.COD_SOFTLAND,a.N_PARTICIPANTES,CONVERT(VARCHAR(10),F.FECHA_VENCIMIENTO,105) AS FECHA_VENCIMIENTO,"
    qsHc=qsHc&"(CASE WHEN a.TIPO_DOC='0' then 'Orden de Compra '+a.ORDEN_COMPRA "
	qsHc=qsHc&"WHEN a.TIPO_DOC='1' then 'Vale Vista '+a.ORDEN_COMPRA "
	qsHc=qsHc&"WHEN a.TIPO_DOC='2' then 'Depósito Cheque '+a.ORDEN_COMPRA "
	qsHc=qsHc&"WHEN a.TIPO_DOC='3' then 'Transferencia '+a.ORDEN_COMPRA "
	qsHc=qsHc&"WHEN a.TIPO_DOC='4' then 'Carta Compromiso '+a.ORDEN_COMPRA "
	qsHc=qsHc&"END) as 'Tipo Documento',f.MONTO,f.VALOR_CURSO FROM FACTURAS F "
	qsHc=qsHc&" inner join AUTORIZACION a on a.ID_AUTORIZACION=f.ID_AUTORIZACION "
	qsHc=qsHc&" inner join empresas e on e.ID_EMPRESA=f.ID_EMPRESA "
	qsHc=qsHc&" inner join programa p on p.ID_PROGRAMA=a.ID_PROGRAMA "
	qsHc=qsHc&" inner join CURRICULO c on c.ID_MUTUAL=p.ID_MUTUAL "
	qsHc=qsHc&" WHERE F.FECHA_EMISION>=CONVERT(date, '"&request("i")&"') and  "
	qsHc=qsHc&" F.FECHA_EMISION<=CONVERT(date, '"&request("t")&"') "
	qsHc=qsHc&" order by f.FACTURA asc"

	'qsHc=qsHc&" and c.ID_MUTUAL=29"

	set rsHc =  conn.execute(qsHc)%>

    <table width="4000" border="1">
    <!--<tr>
      <td width="100" align="center"><b><font size="2">Mes</font></b></td>
      <td width="50" align="center"><b><font size="2">Año</font></b></td>
      <td width="100" align="center"><b><font size="2">Fecha de Inicio</font></b></td>
      <td width="120" align="center"><b><font size="2">Fecha de T&eacute;rmino</font></b></td>
      <td width="400" align="center"><b><font size="2">Sala / Sede</font></b></td>
      <td width="250" align="center"><b><font size="2">Relator</font></b></td>
      <td width="100" align="center"><b><font size="2">C&oacute;digo de Curso</font></b></td> 
      <td width="550" align="center"><b><font size="2">Nombre del Curso</font></b></td>      
      <td width="100" align="center"><b><font size="2">Rut de Empresa</font></b></td>
      <td width="350" align="center"><b><font size="2">Razón Social</font></b></td>
      <td width="150" align="center"><b><font size="2">ID/Rut Trabajador</font></b></td>
      <td width="330" align="center"><b><font size="2">Nombre del Trabajador</font></b></td>
      <td width="80" align="center"><b><font size="2">Calificaci&oacute;n</font></b></td>
      <td width="80" align="center"><b><font size="2">Asistencia</font></b></td>
      <td width="80" align="center"><b><font size="2">Evaluaci&oacute;n</font></b></td>  
      <td width="150" align="center"><b><font size="2">Estado</font></b></td>           
    </tr>-->
	<%
    while not rsHc.eof
    %>
    <tr>
      <td align="center"><%=rsHc("tipo_doc")%></td>
	  <td align="center"><%=rsHc("sub_tipo")%></td> 
	  <td align="center"><%=rsHc("estado_doc")%></td>  
	  <td align="center"><%=rsHc("bodega")%></td>  
	  <td align="center"><%=rsHc("FACTURA")%></td> 
	  <td align="center">&nbsp;<%=replace(cstr(rsHc("FECHA_FACTURA")),"-","/")%></td>  
	  <td align="center">&nbsp;<%=cstr(rsHc("conc_doc"))%></td>  
	  <td align="center">&nbsp;<%=replace(cstr(rsHc("FECHA_VENCIMIENTO")),"-","/")%></td> 
	  <td align="right"><%=cstr(rsHc("Tipo Documento"))%></td>  
	  <td align="center"><%=rsHc("conc_doc")%></td>   
      
      <td align="center">0</td>
	  <td align="center">0</td> 
	  <td align="center">0</td>  
	  <td align="center">0</td>  
	  <td align="center">0</td> 
	  <td align="center">0</td>  
	  <td align="center"><%=mid(rsHc("RUT"), 1,len(rsHc("RUT"))-2)%></td>  
	  <td align="center"></td> 
	  <td align="center"></td>  
	  <td align="center"></td>   
      
      <td align="center"></td>
	  <td align="center"></td> 
	  <td align="center"></td>  
	  <td align="center"></td>  
	  <td align="center"></td> 
	  <td align="center"></td>  
	  <td align="center"></td>  
	  <td align="center"></td> 
	  <td align="center"></td>  
	  <td align="center"></td>   
      
      <td align="center"></td>
	  <td align="center">0</td> 
	  <td align="center">0</td>  
	  <td align="center">0</td>  
	  <td align="center">0</td> 
	  <td align="center">0</td>  
	  <td align="center">0</td>  
	  <td align="center">0</td> 
	  <td align="center">0</td>  
	  <td align="center">0</td>   
      
      <td align="center">0</td>
	  <td align="center">0</td> 
	  <td align="center">0</td>  
	  <td align="center"></td>  
      <%if(rsHc("estado_doc")<>"N")then%>
	  		<td align="center"><%=rsHc("MONTO")%></td> 
      <%else%>
      		<td align="center">0</td>
      <%end if%>
	  <td align="center"><%=rsHc("COD_SOFTLAND")%></td>  
	  <td align="center"></td>
	  <td align="center"></td>
	  <td align="center">1</td>
	  <td align="center">0</td>   
      <%if(rsHc("estado_doc")<>"N")then%>   
              <td align="center"><%=rsHc("MONTO")%></td>
              <td align="center"><%=rsHc("MONTO")%></td> 
      <%else%>
      		  <td align="center">0</td>
			  <td align="center">0</td>            
      <%end if%>
	  <td align="center">0</td>  
	  <td align="center">0</td>  
	  <td align="center">0</td> 
	  <td align="center">0</td>  
	  <td align="center">0</td>  
	  <td align="center">0</td> 
	  <td align="center">0</td>  
	  <td align="center">0</td>   

      <td align="center">0</td>
	  <td align="center">0</td> 
	  <td align="center">0</td>  
	  <td align="center">0</td>  
	  <td align="center"></td> 
	  <td align="center"></td>  
	  <td align="center"></td>  
	  <td align="center"></td> 
	  <td align="center"></td>  
    </tr>	
    <%
        rsHc.MoveNext
    wend
    %>
    </table>
