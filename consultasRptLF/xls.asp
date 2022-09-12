<!--#include file="../conexion.asp"-->
<%
fecha = year(now)&right("0"&month(now()),2)&right("0"&day(now()),2)
fecha = fecha&right("0"&hour(now()),2)&right("0"&minute(now()),2)&right("0"&second(now()),2)
Response.ContentType ="application/vnd.ms-excel"

nomArch=""
if(request("insc")="1")then
	nomArch="Ventas_Norte_"
elseif(request("insc")="2")then
	nomArch="Pendiente_Liberar_Norte_"
elseif(request("insc")="3")then 
	nomArch="Detalle_Facturación_Norte_"	
end if

Response.AddHeader "content-disposition", "inline; filename="&nomArch&fecha&".xls"
	
	mesSel=""
	if(request("mes")<>"0")then
		mesSel=" Month(PROGRAMA.FECHA_INICIO_)="&request("mes")
	end if

	anoSel=""
	if(request("ano")<>"0")then
		anoSel=" and Year(PROGRAMA.FECHA_INICIO_)="&request("ano")
	else
		anoSel=" and Year(PROGRAMA.FECHA_INICIO_)>2012"
	end if
	
	if(request("insc")="1")then
	
	totVenta=0
	
	qsHc="select EMPRESAS.RUT,UPPER (EMPRESAS.R_SOCIAL)as 'RAZÓN SOCIAL',"
	qsHc=qsHc&"Case Month(PROGRAMA.FECHA_INICIO_) "
	qsHc=qsHc&"When 1 then 'Enero' "
	qsHc=qsHc&"When 2 then 'Febrero' "
	qsHc=qsHc&"When 3 then 'Marzo' "
	qsHc=qsHc&"When 4 then 'Abril' "
	qsHc=qsHc&"When 5 then 'Mayo' "
	qsHc=qsHc&"When 6 then 'Junio' "
	qsHc=qsHc&"When 7 then 'Julio' "
	qsHc=qsHc&"When 8 then 'Agosto' "
	qsHc=qsHc&"When 9 then 'Septiembre' "
	qsHc=qsHc&"When 10 then 'Octubre' "
	qsHc=qsHc&"When 11 then 'Noviembre' "
	qsHc=qsHc&"When 12 then 'Diciembre' End as Mes,Year(PROGRAMA.FECHA_INICIO_) as 'Año', "
	qsHc=qsHc&"AUTORIZACION.VALOR_OC,(CASE WHEN AUTORIZACION.TIPO_DOC='0' then 'Orden de Compra'"
	qsHc=qsHc&" WHEN AUTORIZACION.TIPO_DOC='1' then 'Vale Vista' "
	qsHc=qsHc&" WHEN AUTORIZACION.TIPO_DOC='2' then 'Depósito Cheque' "
	qsHc=qsHc&" WHEN AUTORIZACION.TIPO_DOC='3' then 'Transferencia' "
	qsHc=qsHc&" WHEN AUTORIZACION.TIPO_DOC='4' then 'Carta Compromiso' " 
	qsHc=qsHc&" END) as 'Tipo Documento',AUTORIZACION.ORDEN_COMPRA, curriculo.CODIGO, CURRICULO.DESCRIPCION,"
	qsHc=qsHc&"(INSTRUCTOR_RELATOR.NOMBRES+' '+INSTRUCTOR_RELATOR.A_PATERNO+' '+INSTRUCTOR_RELATOR.A_MATERNO) as instructor,"
	qsHc=qsHc&"CONVERT(VARCHAR(10),PROGRAMA.FECHA_INICIO_, 105) as FECHA_INICIO_,"
	qsHc=qsHc&"CONVERT(VARCHAR(10),PROGRAMA.FECHA_TERMINO, 105) as FECHA_TERMINO,"
	qsHc=qsHc&"AUTORIZACION.N_PARTICIPANTES,(select count(*) from HISTORICO_CURSOS where "
	qsHc=qsHc&"HISTORICO_CURSOS.ID_AUTORIZACION=AUTORIZACION.ID_AUTORIZACION) as total,"
	qsHc=qsHc&"'http://norte.otecmutual.cl/ordenes/'+ AUTORIZACION.DOCUMENTO_COMPROMISO as DOCUMENTO_COMPROMISO,"
	qsHc=qsHc&"(case AUTORIZACION.ESTADO when 0 then 'Liberada' when 1 then 'Pendiente de Liberar' end) as 'estInsc', "
	qsHc=qsHc&"(case AUTORIZACION.CON_FRANQUICIA when 0 then 'Sin Franquicia' when 1 then 'Con Franquicia' end) as 'Franquicia',"
    qsHc=qsHc&"(case AUTORIZACION.CON_OTIC when 0 then 'Sin OTIC' when 1 then 'Con OTIC' end) as 'OTIC'"
	qsHc=qsHc&",AUTORIZACION.ID_PROGRAMA,AUTORIZACION.ID_BLOQUE,AUTORIZACION.TIPO_DOC,CURRICULO.ID_CECO from AUTORIZACION "
	qsHc=qsHc&" inner join PROGRAMA on PROGRAMA.ID_PROGRAMA=AUTORIZACION.ID_PROGRAMA "
	qsHc=qsHc&" inner join CURRICULO on CURRICULO.ID_MUTUAL=PROGRAMA.ID_MUTUAL "
	qsHc=qsHc&" inner join EMPRESAS on EMPRESAS.ID_EMPRESA=AUTORIZACION.ID_EMPRESA "
	qsHc=qsHc&" inner join bloque_programacion on bloque_programacion.id_bloque=AUTORIZACION.ID_BLOQUE "
	qsHc=qsHc&" inner join INSTRUCTOR_RELATOR on INSTRUCTOR_RELATOR.ID_INSTRUCTOR=bloque_programacion.id_relator "
	qsHc=qsHc&" where "&mesSel&anoSel
	qsHc=qsHc&" order by PROGRAMA.FECHA_INICIO_ asc "

	set rsHc =  conn.execute(qsHc)%>

    <table width="7000" border="1">
    <tr>
      <td width="100" align="center"><b><font size="2">Rut de Empresa</font></b></td>
      <td width="350" align="center"><b><font size="2">Razón Social</font></b></td>
      <td width="100" align="center"><b><font size="2">Mes</font></b></td>
      <td width="50" align="center"><b><font size="2">Año</font></b></td>      
      <td width="80" align="center"><b><font size="2">Valor</font></b></td>
      <td width="200" align="center"><b><font size="2">Tipo de Documento</font></b></td>
      <td width="150" align="center"><b><font size="2">N° de Documento</font></b></td>      
      <td width="100" align="center"><b><font size="2">C&eacute;digo de Curso</font></b></td> 
       <td width="80" align="center"><b><font size="2">Id Ceco</font></b></td> 
      <td width="200" align="center"><b><font size="2">Nombre del Curso</font></b></td> 
      <!--<td width="250" align="center"><b><font size="2">Relator</font></b></td>-->              
      <td width="100" align="center"><b><font size="2">Fecha de Inicio</font></b></td>
      <td width="120" align="center"><b><font size="2">Fecha de T&eacute;rmino</font></b></td>
      <td width="80" align="center"><b><font size="2">N° Part.</font></b></td>
      <!--<td width="100" align="center"><b><font size="2">Ver Documento</font></b></td> -->  
      <td width="150" align="center"><b><font size="2">Estado</font></b></td>  
      <td width="150" align="center"><b><font size="2">Franquicia</font></b></td> 
      <td width="150" align="center"><b><font size="2">OTIC</font></b></td>               
    </tr>
	<%
    while not rsHc.eof
    %>
    <tr>
      <td align="right"><font size="2"><%=rsHc("RUT")%></font></td>
      <td><font size="2"><%=rsHc("RAZÓN SOCIAL")%></font></td>
      <td align="center"><font size="2"><%=rsHc("Mes")%></font></td>
      <td align="center"><font size="2"><%=rsHc("Año")%></font></td>        
      <td><font size="2"><%=rsHc("VALOR_OC")%></font></td>
      <td><font size="2"><%=rsHc("Tipo Documento")%></font></td>
      <td align="right"><font size="2"><%=rsHc("ORDEN_COMPRA")%></font></td>  
      <td align="right"><font size="2"><%=rsHc("CODIGO")%></font></td>   
      <td align="center"><b><font size="2"><%=rsHc("id_ceco")%></font></b></td>                 
      <td><font size="2"><%=rsHc("DESCRIPCION")%></font></td>
      <!--<td><font size="2"><%=rsHc("instructor")%></font></td>-->
      <td align="center"><font size="2"><%=rsHc("FECHA_INICIO_")%></font></td>
      <td align="center"><font size="2"><%=rsHc("FECHA_TERMINO")%></font></td>
      <td align="center"><font size="2"><%=rsHc("total")%></font></td>
      <!--<td align="center"><font size="2"><a href="<%=rsHc("DOCUMENTO_COMPROMISO")%>">Ver</a></font></td>-->  
      <td align="center"><font size="2"><%=rsHc("estInsc")%></font></td>   
      <td align="center"><font size="2"><%=rsHc("Franquicia")%></font></td> 
      <td align="center"><font size="2"><%=rsHc("OTIC")%></font></td>     
    </tr>	
    <%
		totVenta=totVenta+rsHc("VALOR_OC")
        rsHc.MoveNext
    wend
    %>
    </table>    
    <table border="0">
    <tr>
      <td colspan="4">&nbsp;</td>
      <td><b><%=totVenta%></b></td>
      <td colspan="10">&nbsp;</td>
    </tr>	    
    </table>
    <%elseif(request("insc")="2")then
	subtotPend=0
	totPend=0
	
	qsHc="select EMPRESAS.RUT,UPPER (EMPRESAS.R_SOCIAL)as 'RAZÓN SOCIAL',"
	qsHc=qsHc&"Case Month(PROGRAMA.FECHA_INICIO_) "
	qsHc=qsHc&"When 1 then 'Enero' "
	qsHc=qsHc&"When 2 then 'Febrero' "
	qsHc=qsHc&"When 3 then 'Marzo' "
	qsHc=qsHc&"When 4 then 'Abril' "
	qsHc=qsHc&"When 5 then 'Mayo' "
	qsHc=qsHc&"When 6 then 'Junio' "
	qsHc=qsHc&"When 7 then 'Julio' "
	qsHc=qsHc&"When 8 then 'Agosto' "
	qsHc=qsHc&"When 9 then 'Septiembre' "
	qsHc=qsHc&"When 10 then 'Octubre' "
	qsHc=qsHc&"When 11 then 'Noviembre' "
	qsHc=qsHc&"When 12 then 'Diciembre' End as Mes,Year(PROGRAMA.FECHA_INICIO_) as 'Año', "
	qsHc=qsHc&"AUTORIZACION.VALOR_OC,(CASE WHEN AUTORIZACION.TIPO_DOC='0' then 'Orden de Compra'"
	qsHc=qsHc&" WHEN AUTORIZACION.TIPO_DOC='1' then 'Vale Vista' "
	qsHc=qsHc&" WHEN AUTORIZACION.TIPO_DOC='2' then 'Depósito Cheque' "
	qsHc=qsHc&" WHEN AUTORIZACION.TIPO_DOC='3' then 'Transferencia' "
	qsHc=qsHc&" WHEN AUTORIZACION.TIPO_DOC='4' then 'Carta Compromiso' " 
	qsHc=qsHc&" END) as 'Tipo Documento',AUTORIZACION.ORDEN_COMPRA, curriculo.CODIGO, CURRICULO.DESCRIPCION,"
	qsHc=qsHc&"(INSTRUCTOR_RELATOR.NOMBRES+' '+INSTRUCTOR_RELATOR.A_PATERNO+' '+INSTRUCTOR_RELATOR.A_MATERNO) as instructor,"
	qsHc=qsHc&"CONVERT(VARCHAR(10),PROGRAMA.FECHA_INICIO_, 105) as FECHA_INICIO_,"
	qsHc=qsHc&"CONVERT(VARCHAR(10),PROGRAMA.FECHA_TERMINO, 105) as FECHA_TERMINO,"
	qsHc=qsHc&"AUTORIZACION.N_PARTICIPANTES,(select count(*) from HISTORICO_CURSOS where "
	qsHc=qsHc&"HISTORICO_CURSOS.ID_AUTORIZACION=AUTORIZACION.ID_AUTORIZACION) as total,"
	qsHc=qsHc&"'http://norte.otecmutual.cl/ordenes/'+ AUTORIZACION.DOCUMENTO_COMPROMISO as DOCUMENTO_COMPROMISO,"
	qsHc=qsHc&"(case AUTORIZACION.ESTADO when 0 then 'Liberada' when 1 then 'Pendiente de Liberar' end) as 'estInsc', "
	qsHc=qsHc&"(case AUTORIZACION.CON_FRANQUICIA when 0 then 'Sin Franquicia' when 1 then 'Con Franquicia' end) as 'Franquicia',"
    qsHc=qsHc&"(case AUTORIZACION.CON_OTIC when 0 then 'Sin OTIC' when 1 then 'Con OTIC' end) as 'OTIC'"	
	qsHc=qsHc&",AUTORIZACION.ID_PROGRAMA,AUTORIZACION.ID_BLOQUE,AUTORIZACION.TIPO_DOC,CURRICULO.ID_CECO from AUTORIZACION "
	qsHc=qsHc&" inner join PROGRAMA on PROGRAMA.ID_PROGRAMA=AUTORIZACION.ID_PROGRAMA "
	qsHc=qsHc&" inner join CURRICULO on CURRICULO.ID_MUTUAL=PROGRAMA.ID_MUTUAL "
	qsHc=qsHc&" inner join EMPRESAS on EMPRESAS.ID_EMPRESA=AUTORIZACION.ID_EMPRESA "
	qsHc=qsHc&" inner join bloque_programacion on bloque_programacion.id_bloque=AUTORIZACION.ID_BLOQUE "
	qsHc=qsHc&" inner join INSTRUCTOR_RELATOR on INSTRUCTOR_RELATOR.ID_INSTRUCTOR=bloque_programacion.id_relator "
	qsHc=qsHc&" where AUTORIZACION.estado=1"
    qsHc=qsHc&" and PROGRAMA.FECHA_INICIO_<convert(date,DATEADD(month,1,'01-"&request("mes")&"-"&request("ano")&"'))"
    qsHc=qsHc&" and PROGRAMA.FECHA_INICIO_>=convert(date,'01-01-2013')"	
	qsHc=qsHc&" order by PROGRAMA.FECHA_INICIO_ asc "

	set rsHc =  conn.execute(qsHc)%>

    <table width="7000" border="1">
    <tr>
      <td width="100" align="center"><b><font size="2">Rut de Empresa</font></b></td>
      <td width="350" align="center"><b><font size="2">Razón Social</font></b></td>
      <td width="100" align="center"><b><font size="2">Mes</font></b></td>
      <td width="50" align="center"><b><font size="2">Año</font></b></td>      
      <td width="80" align="center"><b><font size="2">Valor</font></b></td>
      <td width="200" align="center"><b><font size="2">Tipo de Documento</font></b></td>
      <td width="150" align="center"><b><font size="2">N° de Documento</font></b></td>      
      <td width="100" align="center"><b><font size="2">C&eacute;digo de Curso</font></b></td> 
            <td width="80" align="center"><b><font size="2">Id Ceco</font></b></td> 
      <td width="750" align="center"><b><font size="2">Nombre del Curso</font></b></td> 
      <!--<td width="250" align="center"><b><font size="2">Relator</font></b></td>     -->         
      <td width="100" align="center"><b><font size="2">Fecha de Inicio</font></b></td>
      <td width="120" align="center"><b><font size="2">Fecha de T&eacute;rmino</font></b></td>
      <td width="80" align="center"><b><font size="2">N° Part.</font></b></td>
      <!--<td width="100" align="center"><b><font size="2">Ver Documento</font></b></td>-->   
      <td width="150" align="center"><b><font size="2">Estado</font></b></td>
      <td width="150" align="center"><b><font size="2">Franquicia</font></b></td> 
      <td width="150" align="center"><b><font size="2">OTIC</font></b></td>                        
    </tr>
	<%
    while not rsHc.eof
    %>
    <tr>
      <td align="right"><font size="2"><%=rsHc("RUT")%></font></td>
      <td><font size="2"><%=rsHc("RAZÓN SOCIAL")%></font></td>
      <td align="center"><font size="2"><%=rsHc("Mes")%></font></td>
      <td align="center"><font size="2"><%=rsHc("Año")%></font></td>       
      <td><font size="2"><%=rsHc("VALOR_OC")%></font></td>
      <td><font size="2"><%=rsHc("Tipo Documento")%></font></td>
      <td align="right"><font size="2"><%=rsHc("ORDEN_COMPRA")%></font></td>  
      <td align="right"><font size="2"><%=rsHc("CODIGO")%></font></td>  
      <td align="center"><b><font size="2"><%=rsHc("id_ceco")%></font></b></td>                  
      <td><font size="2"><%=rsHc("DESCRIPCION")%></font></td>
      <!--<td><font size="2"><%=rsHc("instructor")%></font></td>-->
      <td align="center"><font size="2"><%=rsHc("FECHA_INICIO_")%></font></td>
      <td align="center"><font size="2"><%=rsHc("FECHA_TERMINO")%></font></td>
      <td align="center"><font size="2"><%=rsHc("total")%></font></td>
      <!--<td align="center"><font size="2"><a href="<%=rsHc("DOCUMENTO_COMPROMISO")%>">Ver</a></font></td> --> 
      <td align="center"><font size="2"><%=rsHc("estInsc")%></font></td>    
      <td align="center"><font size="2"><%=rsHc("Franquicia")%></font></td> 
      <td align="center"><font size="2"><%=rsHc("OTIC")%></font></td>             
    </tr>	
    <%
		subtotPend=subtotPend+rsHc("VALOR_OC")
        rsHc.MoveNext
    wend

	totPend=totPend+subtotPend
	
    %>
    </table>    
    <table border="0">
    <tr>
      <td colspan="4">&nbsp;</td>
      <td><b><%=subtotPend%></b></td>
      <td colspan="10">&nbsp;</td>
    </tr>	
    <tr>
      <td colspan="15">&nbsp;</td>
    </tr>      
    <tr>
      <td colspan="15"><b>Liberadas No Facturadas</b></td>
    </tr>	       
    </table>
    <% 
	subtotPend=0
	qsHc2="select EMPRESAS.RUT,UPPER (EMPRESAS.R_SOCIAL)as 'RAZÓN SOCIAL',"
	qsHc2=qsHc2&"Case Month(PROGRAMA.FECHA_INICIO_) "
	qsHc2=qsHc2&"When 1 then 'Enero' "
	qsHc2=qsHc2&"When 2 then 'Febrero' "
	qsHc2=qsHc2&"When 3 then 'Marzo' "
	qsHc2=qsHc2&"When 4 then 'Abril' "
	qsHc2=qsHc2&"When 5 then 'Mayo' "
	qsHc2=qsHc2&"When 6 then 'Junio' "
	qsHc2=qsHc2&"When 7 then 'Julio' "
	qsHc2=qsHc2&"When 8 then 'Agosto' "
	qsHc2=qsHc2&"When 9 then 'Septiembre' "
	qsHc2=qsHc2&"When 10 then 'Octubre' "
	qsHc2=qsHc2&"When 11 then 'Noviembre' "
	qsHc2=qsHc2&"When 12 then 'Diciembre' End as Mes,Year(PROGRAMA.FECHA_INICIO_) as 'Año', "
	qsHc2=qsHc2&"AUTORIZACION.VALOR_OC,(CASE WHEN AUTORIZACION.TIPO_DOC='0' then 'Orden de Compra'"
	qsHc2=qsHc2&" WHEN AUTORIZACION.TIPO_DOC='1' then 'Vale Vista' "
	qsHc2=qsHc2&" WHEN AUTORIZACION.TIPO_DOC='2' then 'Depósito Cheque' "
	qsHc2=qsHc2&" WHEN AUTORIZACION.TIPO_DOC='3' then 'Transferencia' "
	qsHc2=qsHc2&" WHEN AUTORIZACION.TIPO_DOC='4' then 'Carta Compromiso' " 
	qsHc2=qsHc2&" END) as 'Tipo Documento',AUTORIZACION.ORDEN_COMPRA, curriculo.CODIGO, CURRICULO.DESCRIPCION,"
	qsHc2=qsHc2&"(INSTRUCTOR_RELATOR.NOMBRES+' '+INSTRUCTOR_RELATOR.A_PATERNO+' '+INSTRUCTOR_RELATOR.A_MATERNO) as instructor,"
	qsHc2=qsHc2&"CONVERT(VARCHAR(10),PROGRAMA.FECHA_INICIO_, 105) as FECHA_INICIO_,"
	qsHc2=qsHc2&"CONVERT(VARCHAR(10),PROGRAMA.FECHA_TERMINO, 105) as FECHA_TERMINO,"
	qsHc2=qsHc2&"AUTORIZACION.N_PARTICIPANTES,(select count(*) from HISTORICO_CURSOS where "
	qsHc2=qsHc2&"HISTORICO_CURSOS.ID_AUTORIZACION=AUTORIZACION.ID_AUTORIZACION) as total,"
	qsHc2=qsHc2&"'http://norte.otecmutual.cl/ordenes/'+ AUTORIZACION.DOCUMENTO_COMPROMISO as DOCUMENTO_COMPROMISO,"
	qsHc2=qsHc2&"(case AUTORIZACION.ESTADO when 0 then 'Liberada' when 1 then 'Pendiente de Liberar' end) as 'estInsc', "
	qsHc2=qsHc2&"(case AUTORIZACION.CON_FRANQUICIA when 0 then 'Sin Franquicia' when 1 then 'Con Franquicia' end) as 'Franquicia',"
    qsHc2=qsHc2&"(case AUTORIZACION.CON_OTIC when 0 then 'Sin OTIC' when 1 then 'Con OTIC' end) as 'OTIC'"	
   qsHc2=qsHc2&",AUTORIZACION.ID_PROGRAMA,AUTORIZACION.ID_BLOQUE,AUTORIZACION.TIPO_DOC,CURRICULO.ID_CECO from AUTORIZACION "
	qsHc2=qsHc2&" inner join PROGRAMA on PROGRAMA.ID_PROGRAMA=AUTORIZACION.ID_PROGRAMA "
	qsHc2=qsHc2&" inner join CURRICULO on CURRICULO.ID_MUTUAL=PROGRAMA.ID_MUTUAL "
	qsHc2=qsHc2&" inner join EMPRESAS on EMPRESAS.ID_EMPRESA=AUTORIZACION.ID_EMPRESA "
	qsHc2=qsHc2&" inner join bloque_programacion on bloque_programacion.id_bloque=AUTORIZACION.ID_BLOQUE "
	qsHc2=qsHc2&" inner join INSTRUCTOR_RELATOR on INSTRUCTOR_RELATOR.ID_INSTRUCTOR=bloque_programacion.id_relator "
	qsHc2=qsHc2&" where AUTORIZACION.estado=0 and AUTORIZACION.FACTURADO=1"
    qsHc2=qsHc2&" and PROGRAMA.FECHA_INICIO_<convert(date,DATEADD(month,1,'01-"&request("mes")&"-"&request("ano")&"'))"
    qsHc2=qsHc2&" and PROGRAMA.FECHA_INICIO_>=convert(date,'01-01-2013') and autorizacion.id_autorizacion not in (42574,42518,49777)"	
	qsHc2=qsHc2&" order by PROGRAMA.FECHA_INICIO_ asc "

	set rsHc2 =  conn.execute(qsHc2)%>

    <table border="1">
    <tr>
      <td width="100" align="center"><b><font size="2">Rut de Empresa</font></b></td>
      <td width="350" align="center"><b><font size="2">Razón Social</font></b></td>
      <td width="100" align="center"><b><font size="2">Mes</font></b></td>
      <td width="50" align="center"><b><font size="2">Año</font></b></td>      
      <td width="80" align="center"><b><font size="2">Valor</font></b></td>
      <td width="200" align="center"><b><font size="2">Tipo de Documento</font></b></td>
      <td width="150" align="center"><b><font size="2">N° de Documento</font></b></td>      
      <td width="100" align="center"><b><font size="2">C&eacute;digo de Curso</font></b></td> 
            <td width="80" align="center"><b><font size="2">Id Ceco</font></b></td> 
      <td width="750" align="center"><b><font size="2">Nombre del Curso</font></b></td> 
      <!--<td width="250" align="center"><b><font size="2">Relator</font></b></td> -->             
      <td width="100" align="center"><b><font size="2">Fecha de Inicio</font></b></td>
      <td width="120" align="center"><b><font size="2">Fecha de T&eacute;rmino</font></b></td>
      <td width="80" align="center"><b><font size="2">N° Part.</font></b></td>
      <!--<td width="100" align="center"><b><font size="2">Ver Documento</font></b></td> -->  
      <td width="150" align="center"><b><font size="2">Estado</font></b></td>
      <td width="150" align="center"><b><font size="2">Franquicia</font></b></td> 
      <td width="150" align="center"><b><font size="2">OTIC</font></b></td>                        
    </tr>
	<%
    while not rsHc2.eof
    %>
    <tr>
      <td align="right"><font size="2"><%=rsHc2("RUT")%></font></td>
      <td><font size="2"><%=rsHc2("RAZÓN SOCIAL")%></font></td>
      <td align="center"><font size="2"><%=rsHc2("Mes")%></font></td>
      <td align="center"><font size="2"><%=rsHc2("Año")%></font></td>       
      <td><font size="2"><%=rsHc2("VALOR_OC")%></font></td>
      <td><font size="2"><%=rsHc2("Tipo Documento")%></font></td>
      <td align="right"><font size="2"><%=rsHc2("ORDEN_COMPRA")%></font></td>  
      <td align="right"><font size="2"><%=rsHc2("CODIGO")%></font></td>        
      <td align="center"><b><font size="2"><%=rsHc2("id_ceco")%></font></b></td>            
      <td><font size="2"><%=rsHc2("DESCRIPCION")%></font></td>
      <!--<td><font size="2"><%=rsHc2("instructor")%></font></td>-->
      <td align="center"><font size="2"><%=rsHc2("FECHA_INICIO_")%></font></td>
      <td align="center"><font size="2"><%=rsHc2("FECHA_TERMINO")%></font></td>
      <td align="center"><font size="2"><%=rsHc2("total")%></font></td>
      <!--<td align="center"><font size="2"><a href="<%=rsHc2("DOCUMENTO_COMPROMISO")%>">Ver</a></font></td> --> 
      <td align="center"><font size="2"><%=rsHc2("estInsc")%></font></td>   
      <td align="center"><font size="2"><%=rsHc2("Franquicia")%></font></td> 
      <td align="center"><font size="2"><%=rsHc2("OTIC")%></font></td>           
    </tr>	
    <%
		subtotPend=subtotPend+rsHc2("VALOR_OC")
        rsHc2.MoveNext
    wend
	
	totPend=totPend+subtotPend
    %>
    </table>    
    <table border="0">
    <tr>
      <td colspan="4">&nbsp;</td>
      <td><b><%=subtotPend%></b></td>
      <td colspan="10">&nbsp;</td>
    </tr>	
    <tr>
      <td colspan="15">&nbsp;</td>
    </tr>	
    <tr>
      <td colspan="3">&nbsp;</td>
      <td>Total:</td>
      <td><b><%=totPend%></b></td>
      <td colspan="10">&nbsp;</td>
    </tr>	        
    </table>   
    <%elseif(request("insc")="3")then  
	subtotPend=0
	
	qsHc="select u.RUT,USUARIO=DBO.MayMinTexto(u.NOMBRES+' '+u.A_PATERNO+' '+u.A_MATERNO),"&_
		 " Total=(select count(*) from facturas f2 WHERE YEAR(f2.FECHA_EMISION)="&request("ano")&" and MONTH(f2.FECHA_EMISION)="&request("mes")&_
		 " and f2.FACTURADO=u.ID_USUARIO),u.ID_USUARIO from usuarios u "&_
		 " where u.ID_USUARIO in (select f.FACTURADO from facturas f "&_
		 " WHERE YEAR(f.FECHA_EMISION)="&request("ano")&" and MONTH(f.FECHA_EMISION)="&request("mes")&") order BY u.NOMBRES"

	set rsHc =  conn.execute(qsHc)%>

    <table width="7000" border="1">
    <tr>
      <td colspan="3"><b><font size="2">Resumen Facturaci&oacute;n</font></b></td>
    </tr>    
    <tr>
      <td width="100" align="center"><b><font size="2">Rut Usuario</font></b></td>
      <td width="350" align="center"><b><font size="2">Usuario</font></b></td>
      <td width="100" align="center"><b><font size="2">Total</font></b></td>
    </tr>
	<%
    while not rsHc.eof
    %>
    <tr>
      <td align="right"><font size="2"><%=rsHc("RUT")%></font></td>
      <td><font size="2"><%=rsHc("USUARIO")%></font></td>
      <td><font size="2"><%=rsHc("Total")%></font></td>
    </tr>	
    <%
		subtotPend=subtotPend+rsHc("Total")
        rsHc.MoveNext
    wend
	
	rsHc.MoveFirst
    %>
    </table>    
    <table border="0">
    <tr>
      <td colspan="2">&nbsp;</td>
      <td><b><%=subtotPend%></b></td>
    </tr>	
    <tr>
      <td colspan="3">&nbsp;</td>
    </tr>      
    <tr>
      <td colspan="3"><b><font size="2">Detalle Facturación</font></b></td>
    </tr>	       
    </table>
    <%
    while not rsHc.eof
    %>
	<table border="0">
    <tr>
      <td colspan="2">&nbsp;</td>
    </tr>
    <tr>
      <td colspan="2"><b>Facturado por <%=rsHc("USUARIO")%></b></td>
    </tr>
    <tr>
      <td colspan="2">&nbsp;</td>
    </tr>        
    </table>     
    <% 
	qsHc2="SELECT E.RUT as rut,dbo.MayMinTexto (E.R_SOCIAL) as empresa, "&_
		  " CONVERT(VARCHAR(10),P.FECHA_INICIO_, 105) as FECHA_INICIO, "&_
		  " CONVERT(VARCHAR(10),P.FECHA_TERMINO, 105) as FECHA_TERMINO, "&_
		  " C.CODIGO, dbo.MayMinTexto (C.NOMBRE_CURSO) as nombre,/*A.N_PARTICIPANTES,*/  "&_
		  " CONVERT(VARCHAR(10),F.FECHA_EMISION,105) as FECHA_EMISION,F.FACTURA, "&_
		  " (CASE F.ESTADO WHEN 1 THEN 'Vigente' WHEN 0 THEN 'Anulada' WHEN 2 THEN 'Vig/Ref' end) as 'DESCEST', "&_
		  " f.MONTO,a.TIPO_DOC,a.ORDEN_COMPRA, "&_
		  " (CASE a.CON_FRANQUICIA WHEN 1 THEN 'Si' ELSE 'No' end) as 'FRANQUICIA', "&_
		  " (CASE a.CON_OTIC WHEN 1 THEN 'Si' ELSE 'No' end) as 'OTIC',A.n_reg_sence, "&_
		  " a.FECHA_REV_CIERRE,f.FECHA_FACTURA,a.FECHA__AUTORIZACION, "&_
		  " USUARIO=DBO.MayMinTexto(u.NOMBRES+' '+u.A_PATERNO+' '+u.A_MATERNO), "&_
		  "(CASE WHEN A.TIPO_DOC='0' then 'Orden de Compra' "&_
		  " WHEN A.TIPO_DOC='1' then 'Vale Vista' "&_
		  " WHEN A.TIPO_DOC='2' then 'Depósito Cheque' "&_
		  " WHEN A.TIPO_DOC='3' then 'Transferencia' "&_
		  " WHEN A.TIPO_DOC='4' then 'Carta Compromiso' " &_
		  " END) as 'Tipo Documento',A.ORDEN_COMPRA,c.ID_CECO "&_		 
		  " FROM FACTURAS F  "&_
		  "  inner join AUTORIZACION A on A.ID_AUTORIZACION=F.ID_AUTORIZACION  "&_
		  "  inner join EMPRESAS EI ON EI.ID_EMPRESA=A.ID_EMPRESA  "&_
		  "  inner join EMPRESAS E on E.ID_EMPRESA=F.ID_EMPRESA  "&_
		  "  inner join PROGRAMA P on P.ID_PROGRAMA=A.ID_PROGRAMA    "&_
		  "  inner join CURRICULO C on C.ID_MUTUAL=P.ID_MUTUAL   "&_
		  "  inner join bloque_programacion BQ on BQ.id_bloque=A.ID_BLOQUE  "&_
		  "  inner join usuarios u on u.ID_USUARIO=f.FACTURADO  "&_
		  "  WHERE YEAR(f.FECHA_EMISION)="&request("ano")&" and MONTH(f.FECHA_EMISION)="&request("mes")&_
		  " AND u.ID_USUARIO="&rsHc("ID_USUARIO")&" order by f.FACTURA"

	set rsHc2 =  conn.execute(qsHc2)%>

    <table border="1">
    <tr>
      <td width="100" align="center"><b><font size="2">Rut de Empresa</font></b></td>
      <td width="350" align="center"><b><font size="2">Razón Social</font></b></td>
      <td width="100" align="center"><b><font size="2">Fecha de Inicio</font></b></td>
      <td width="120" align="center"><b><font size="2">Fecha de T&eacute;rmino</font></b></td>   
      <td width="100" align="center"><b><font size="2">C&eacute;digo de Curso</font></b></td> 
	  <td width="80" align="center"><b><font size="2">Id Ceco</font></b></td>       
      <td width="750" align="center"><b><font size="2">Nombre del Curso</font></b></td>
	  <td width="150" align="center"><b><font size="2">Tipo Documento</font></b></td>    
	  <td width="150" align="center"><b><font size="2">N Documento</font></b></td>           
      <td width="120" align="center"><b><font size="2">Fecha de Emisi&oacute;n</font></b></td> 
      <td width="120" align="center"><b><font size="2">N&deg; Factura</font></b></td>       
      <td width="150" align="center"><b><font size="2">Estado Factura</font></b></td>
      <td width="150" align="center"><b><font size="2">Monto</font></b></td>
      <td width="150" align="center"><b><font size="2">Franquicia</font></b></td> 
      <td width="150" align="center"><b><font size="2">OTIC</font></b></td>          
	  <td width="150" align="center"><b><font size="2">N&deg; Reg. Sence</font></b></td> 
	  <td width="150" align="center"><b><font size="2">Fecha Autorizaci&oacute;n</font></b></td>         
	  <td width="150" align="center"><b><font size="2">Fecha Rev. Cierre</font></b></td>
	  <td width="150" align="center"><b><font size="2">Fecha Factura</font></b></td>
    </tr>
	<%
    while not rsHc2.eof
    %>
    <tr>
      <td align="right"><font size="2"><%=rsHc2("RUT")%></font></td>
      <td><font size="2"><%=rsHc2("empresa")%></font></td>
      <td align="center"><font size="2"><%=rsHc2("FECHA_INICIO")%></font></td>
      <td align="center"><font size="2"><%=rsHc2("FECHA_TERMINO")%></font></td>      
      <td align="right"><font size="2"><%=rsHc2("CODIGO")%></font></td>           
      <td align="center"><b><font size="2"><%=rsHc2("id_ceco")%></font></b></td>         
      <td><font size="2"><%=rsHc2("nombre")%></font></td>
	  <td><font size="2"><%=rsHc2("Tipo Documento")%></font></td>    
	  <td><font size="2"><%=rsHc2("ORDEN_COMPRA")%></font></td>          
      <td align="center"><font size="2"><%=rsHc2("FECHA_EMISION")%></font></td> 
      <td align="center"><font size="2"><%=rsHc2("FACTURA")%></font></td>       
      <td align="center"><font size="2"><%=rsHc2("DESCEST")%></font></td>   
      <td align="center"><font size="2"><%=rsHc2("MONTO")%></font></td>        
      <td align="center"><font size="2"><%=rsHc2("Franquicia")%></font></td> 
      <td align="center"><font size="2"><%=rsHc2("OTIC")%></font></td>     
      <td align="center"><font size="2"><%=rsHc2("n_reg_sence")%></font></td>   
      <td align="center"><font size="2"><%=rsHc2("FECHA__AUTORIZACION")%></font></td> 
      <td align="center"><font size="2"><%=rsHc2("FECHA_REV_CIERRE")%></font></td> 
      <td align="center"><font size="2"><%=rsHc2("FECHA_FACTURA")%></font></td>               
    </tr>	
    <%
        rsHc2.MoveNext
    wend
    %>
	<table border="0">
    <tr>
      <td colspan="2">&nbsp;</td>
    </tr>
    <tr>
      <td colspan="2">&nbsp;</td>
    </tr>
    </table>
	<%
            rsHc.MoveNext
    wend%>
    </table>    
    <%end if%> 