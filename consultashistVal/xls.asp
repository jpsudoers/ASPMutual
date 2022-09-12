<!--#include file="../conexion.asp"-->
<%
fecha = year(now)&right("0"&month(now()),2)&right("0"&day(now()),2)
fecha = fecha&right("0"&hour(now()),2)&right("0"&minute(now()),2)&right("0"&second(now()),2)
Response.ContentType ="application/vnd.ms-excel"
Response.AddHeader "content-disposition", "inline; filename=Historica_Valorizada_"&fecha&".xls"
	
	mesSel=""
	if(request("mes")<>"0")then
		mesSel=" and Month(PROGRAMA.FECHA_INICIO_)="&request("mes")
	end if

	anoSel=""
	if(request("ano")<>"0")then
		anoSel=" and Year(PROGRAMA.FECHA_INICIO_)="&request("ano")
	else
		anoSel=" and Year(PROGRAMA.FECHA_INICIO_)>2012"
	end if
	
	insSel=""
	if(request("insc")<>"Todas")then
		  insSel=" and HISTORICO_CURSOS.ESTADO="&request("insc")
	else
		  insSel=" and HISTORICO_CURSOS.ESTADO in (0,2)"
	end if
	
	qsHc="select CONVERT(VARCHAR(10),PROGRAMA.FECHA_INICIO_, 105) as FECHA_INICIO_,AUTORIZACION.DOCUMENTO_COMPROMISO as 'OC',"
	qsHc=qsHc&"Case Month(PROGRAMA.FECHA_TERMINO) "
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
	qsHc=qsHc&"CONVERT(VARCHAR(10),PROGRAMA.FECHA_TERMINO, 105) as FECHA_TERMINO,"
	qsHc=qsHc&"case when bloque_programacion.id_sede<>27 then sedes.NOMBRE+' - Piso 2, '+sedes.DIRECCION+', '+sedes.CIUDAD "
	qsHc=qsHc&"when bloque_programacion.id_sede=27 then bloque_programacion.nom_sede end AS sede_curso, "
	qsHc=qsHc&"UPPER(INSTRUCTOR_RELATOR.NOMBRES+' '+INSTRUCTOR_RELATOR.A_PATERNO+' '+INSTRUCTOR_RELATOR.A_MATERNO) as instructor," 
	qsHc=qsHc&"CURRICULO.CODIGO,curriculo.DESCRIPCION, " 
	qsHc=qsHc&"empresas.RUT as 'Rut Empresa',UPPER (empresas.R_SOCIAL) as 'Razón Social', " 
	qsHc=qsHc&"(CASE WHEN trabajador.NACIONALIDAD=0 then trabajador.RUT " 
	qsHc=qsHc&" WHEN trabajador.NACIONALIDAD=1 then trabajador.ID_EXTRANJERO " 
	qsHc=qsHc&" END) as 'Rut Trabajador1',UPPER (trabajador.NOMBRES) as NOMBRES,HISTORICO_CURSOS.CALIFICACION, " 
	qsHc=qsHc&"HISTORICO_CURSOS.ASISTENCIA,HISTORICO_CURSOS.EVALUACION,HISTORICO_CURSOS.ID_BLOQUE," 
	qsHc=qsHc&"AUTORIZACION.VALOR_CURSO,(CASE WHEN AUTORIZACION.TIPO_DOC='0' then 'Orden de Compra'"
	qsHc=qsHc&" WHEN AUTORIZACION.TIPO_DOC='1' then 'Vale Vista' "
	qsHc=qsHc&" WHEN AUTORIZACION.TIPO_DOC='2' then 'Depósito Cheque' "
	qsHc=qsHc&" WHEN AUTORIZACION.TIPO_DOC='3' then 'Transferencia' "
	qsHc=qsHc&" WHEN AUTORIZACION.TIPO_DOC='4' then 'Carta Compromiso' " 
	qsHc=qsHc&" WHEN AUTORIZACION.TIPO_DOC='5' then 'EDP' END) as 'Tipo Documento',AUTORIZACION.ORDEN_COMPRA, " 
	qsHc=qsHc&"(case HISTORICO_CURSOS.ESTADO when 0 then 'Liberada' when 2 then 'Pendiente de Liberar' end) as 'estInsc',"
	qsHc=qsHc&"(case PROGRAMA.TIPO when 2 then 'Curso Abierto' when 1 then 'Curso Cerrado' end) as 'tipo_Curso',"
    qsHc=qsHc&"(CASE WHEN AUTORIZACION.CON_FRANQUICIA='0' then 'No' "
	qsHc=qsHc&"WHEN AUTORIZACION.CON_FRANQUICIA='1' then 'Si' END) as CON_FRANQUICIA,"
    qsHc=qsHc&"(CASE WHEN AUTORIZACION.CON_OTIC='0' then 'No' "
	qsHc=qsHc&"WHEN AUTORIZACION.CON_OTIC='1' then 'Si' END) as CON_OTIC,AUTORIZACION.N_REG_SENCE,"
	qsHc=qsHc&"(CASE WHEN AUTORIZACION.CON_OTIC='1' and AUTORIZACION.ID_OTIC is not null then (select eotic.R_SOCIAL "
	qsHc=qsHc&" from EMPRESAS eotic where eotic.ID_EMPRESA=AUTORIZACION.ID_OTIC) else '' END) as NOM_OTIC,	"
	qsHc=qsHc&"AUTORIZACION.FECHA_REV_CIERRE,CURRICULO.ID_CECO,Factura=[dbo].[DET_FACTURAS_AUTORIZACION](HISTORICO_CURSOS.ID_AUTORIZACION),HISTORICO_CURSOS.ID_PROGRAMA,"
	qsHc=qsHc&"'Tipo Venta'=(case when AUTORIZACION.ID_TIPO_VENTA=1 then 'Venta Directa' when AUTORIZACION.ID_TIPO_VENTA=2 then 'Venta Broker' "
	qsHc=qsHc&" else 'Venta Directa' end),'http://norte.otecmutual.cl/ordenes/'+ AUTORIZACION.DOCUMENTO_COMPROMISO as DOCUMENTO_COMPROMISO from HISTORICO_CURSOS " 
	qsHc=qsHc&" inner join trabajador on trabajador.ID_TRABAJADOR=HISTORICO_CURSOS.ID_TRABAJADOR " 
	qsHc=qsHc&" inner join empresas on EMPRESAS.ID_EMPRESA=HISTORICO_CURSOS.ID_EMPRESA " 
	qsHc=qsHc&" inner join PROGRAMA on programa.ID_PROGRAMA=HISTORICO_CURSOS.ID_PROGRAMA " 
	qsHc=qsHc&" inner join CURRICULO on CURRICULO.ID_MUTUAL=PROGRAMA.ID_MUTUAL " 
	qsHc=qsHc&" inner join bloque_programacion on bloque_programacion.id_bloque=HISTORICO_CURSOS.ID_BLOQUE "   
	qsHc=qsHc&" inner join sedes on sedes.ID_SEDE=bloque_programacion.id_sede "  
	qsHc=qsHc&" inner join INSTRUCTOR_RELATOR on INSTRUCTOR_RELATOR.ID_INSTRUCTOR=bloque_programacion.id_relator " 
	qsHc=qsHc&" inner join AUTORIZACION on AUTORIZACION.ID_AUTORIZACION=HISTORICO_CURSOS.ID_AUTORIZACION "	
	qsHc=qsHc&" where HISTORICO_CURSOS.CALIFICACION is not null "&insSel&mesSel&anoSel 
	qsHc=qsHc&" order by PROGRAMA.FECHA_INICIO_,HISTORICO_CURSOS.ID_BLOQUE asc " 


	set rsHc =  conn.execute(qsHc)%>

    <table width="4000" border="1">
    <tr>
      <td width="100" align="center"><b><font size="2">Mes</font></b></td>
      <td width="50" align="center"><b><font size="2">Año</font></b></td>
      <td width="100" align="center"><b><font size="2">Fecha de Inicio</font></b></td>
      <td width="120" align="center"><b><font size="2">Fecha de T&eacute;rmino</font></b></td>
      <td width="400" align="center"><b><font size="2">Sala / Sede</font></b></td>
      <td width="250" align="center"><b><font size="2">Relator</font></b></td>
      <td width="100" align="center"><b><font size="2">C&oacute;digo de Curso</font></b></td> 
      <td width="550" align="center"><b><font size="2">Nombre del Curso</font></b></td>     
      <td width="100" align="center"><b><font size="2">CDR</font></b></td>  
      <td width="100" align="center"><b><font size="2">Rut de Empresa</font></b></td>
      <td width="350" align="center"><b><font size="2">Razón Social</font></b></td>
      <td width="150" align="center"><b><font size="2">ID/Rut Trabajador</font></b></td>
      <td width="330" align="center"><b><font size="2">Nombre del Trabajador</font></b></td>
      <td width="80" align="center"><b><font size="2">Calificaci&oacute;n</font></b></td>
      <td width="80" align="center"><b><font size="2">Asistencia</font></b></td>
      <td width="80" align="center"><b><font size="2">Evaluaci&oacute;n</font></b></td>   
        <td width="80" align="center"><b><font size="2">Valor Asis.</font></b></td> 
        <td width="150" align="center"><b><font size="2">Tipo de Documento</font></b></td> 
        <td width="200" align="center"><b><font size="2">N° de Documento</font></b></td>   
      <td width="150" align="center"><b><font size="2">Estado</font></b></td> 
      <td width="150" align="center"><b><font size="2">Tipo Curso</font></b></td>   
      <td width="150" align="center"><b><font size="2">Con Franquicia</font></b></td> 
      <td width="150" align="center"><b><font size="2">N&deg; Registro Sence</font></b></td> 
      <td width="150" align="center"><b><font size="2">Con Otic</font></b></td> 
      <td width="450" align="center"><b><font size="2">Nombre Otic</font></b></td>         
      <td width="150" align="center"><b><font size="2">Ver Documento</font></b></td>      
      <td width="150" align="center"><b><font size="2">Fecha Liberaci&oacute;n</font></b></td>   
      <td width="100" align="center"><b><font size="2">Factura</font></b></td> 
      <td width="100" align="center"><b><font size="2">ID de Curso</font></b></td> 
      <td width="180" align="center"><b><font size="2">Tipo de Venta</font></b></td> 
    </tr>
	<%
    while not rsHc.eof
    %>
    <tr>
        <td align="center"><b><font size="2"><%=rsHc("Mes")%></font></b></td>
        <td align="center"><b><font size="2"><%=rsHc("Año")%></font></b></td>          
        <td align="center"><b><font size="2"><%=rsHc("FECHA_INICIO_")%></font></b></td>
        <td align="center"><b><font size="2"><%=rsHc("FECHA_TERMINO")%></font></b></td>
        <td><b><font size="2"><%=rsHc("sede_curso")%></font></b></td>
        <td><b><font size="2"><%=rsHc("instructor")%></font></b></td>
        <td align="right"><b><font size="2"><%=rsHc("CODIGO")%></font></b></td>  
        <td><b><font size="2"><%=rsHc("DESCRIPCION")%></font></b></td>        
        <td><b><font size="2"><%=rsHc("id_ceco")%></font></b></td>       
        <td align="right"><b><font size="2"><%=rsHc("Rut Empresa")%></font></b></td>
        <td><b><font size="2"><%=rsHc("Razón Social")%></font></b></td>
        <td align="right"><b><font size="2"><%=rsHc("Rut Trabajador1")%></font></b></td>
        <td><b><font size="2"><%=rsHc("NOMBRES")%></font></b></td>
        <td align="center"><b><font size="2"><%=rsHc("CALIFICACION")%></font></b></td>
        <td align="center"><b><font size="2"><%=rsHc("ASISTENCIA")%></font></b></td>  
        <td align="center"><b><font size="2"><%=rsHc("EVALUACION")%></font></b></td>   
        <td align="center"><b><font size="2"><%=rsHc("VALOR_CURSO")%></font></b></td>	        
	<td><b><font size="2"><%=rsHc("Tipo Documento")%></font></b></td>                                      
        <td align="right"><b><font size="2"><%=rsHc("ORDEN_COMPRA")%></font></b></td>  
        <td align="center"><b><font size="2"><%=rsHc("estInsc")%></font></b></td>   
  		<td align="center"><b><font size="2"><%=rsHc("tipo_Curso")%></font></b></td>  
      	<td align="center"><b><font size="2"><%=rsHc("CON_FRANQUICIA")%></font></b></td> 
      	<td align="center"><b><font size="2"><%=rsHc("N_REG_SENCE")%></font></b></td>       
      	<td align="center"><b><font size="2"><%=rsHc("CON_OTIC")%></font></b></td>   
      	<td><b><font size="2"><%=rsHc("NOM_OTIC")%></font></b></td>                    
 <td align="center"><b><font size="2"><a href="<%=rsHc("DOCUMENTO_COMPROMISO")%>">Ver</a></font></b></td>  
        <td><b><font size="2"><%=rsHc("FECHA_REV_CIERRE")%></font></b></td>  
        <td><b><font size="2"><%=rsHc("FACTURA")%></font></b></td>  
        <td><b><font size="2"><%=rsHc("ID_PROGRAMA")%></font></b></td>  
        <td><b><font size="2"><%=rsHc("Tipo Venta")%></font></b></td>  
    </tr>	
    <%
        rsHc.MoveNext
    wend
    %>
    </table>
