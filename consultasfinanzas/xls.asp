<!--#include file="../conexion.asp"-->
<%
fecha = year(now)&right("0"&month(now()),2)&right("0"&day(now()),2)
fecha = fecha&right("0"&hour(now()),2)&right("0"&minute(now()),2)&right("0"&second(now()),2)
Response.ContentType ="application/vnd.ms-excel"
Response.AddHeader "content-disposition", "inline; filename=Finanzas_"&fecha&".xls"
	
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
	if(request("insc")="0" or request("insc")="1")then
		  insSel=" AUTORIZACION.estado="&request("insc")
	elseif(request("insc")="2")then
		  insSel=" AUTORIZACION.estado=0 and AUTORIZACION.facturado=1 "
	else
		  insSel=" AUTORIZACION.estado in (0,1)"
	end if
	
	qsHc="select EMPRESAS.RUT,UPPER (EMPRESAS.R_SOCIAL)as 'RAZÓN SOCIAL',"
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
	qsHc=qsHc&"AUTORIZACION.VALOR_OC,(CASE WHEN AUTORIZACION.TIPO_DOC='0' then 'Orden de Compra'"
	qsHc=qsHc&" WHEN AUTORIZACION.TIPO_DOC='1' then 'Vale Vista' "
	qsHc=qsHc&" WHEN AUTORIZACION.TIPO_DOC='2' then 'Depósito Cheque' "
	qsHc=qsHc&" WHEN AUTORIZACION.TIPO_DOC='3' then 'Transferencia' "
	qsHc=qsHc&" WHEN AUTORIZACION.TIPO_DOC='4' then 'Carta Compromiso' " 
	qsHc=qsHc&" WHEN AUTORIZACION.TIPO_DOC='5' then 'EDP' " 
	qsHc=qsHc&" END) as 'Tipo Documento',AUTORIZACION.ORDEN_COMPRA, curriculo.CODIGO, CURRICULO.DESCRIPCION,"
	qsHc=qsHc&"(INSTRUCTOR_RELATOR.NOMBRES+' '+INSTRUCTOR_RELATOR.A_PATERNO+' '+INSTRUCTOR_RELATOR.A_MATERNO) as instructor,"
	qsHc=qsHc&"CF=(CASE WHEN AUTORIZACION.CON_FRANQUICIA=1 THEN 'SI' ELSE 'NO' END),"
	qsHc=qsHc&"CONVERT(VARCHAR(10),PROGRAMA.FECHA_INICIO_, 105) as FECHA_INICIO_,"
	qsHc=qsHc&"CONVERT(VARCHAR(10),PROGRAMA.FECHA_TERMINO, 105) as FECHA_TERMINO,"
	qsHc=qsHc&"AUTORIZACION.N_PARTICIPANTES,(select count(*) from HISTORICO_CURSOS where "
	qsHc=qsHc&"HISTORICO_CURSOS.ID_AUTORIZACION=AUTORIZACION.ID_AUTORIZACION) as total,"
	qsHc=qsHc&"'http://norte.otecmutual.cl/ordenes/'+ AUTORIZACION.DOCUMENTO_COMPROMISO as DOCUMENTO_COMPROMISO,"
	qsHc=qsHc&"(case AUTORIZACION.ESTADO when 0 then 'Liberada' when 1 then 'Pendiente de Liberar' end) as 'estInsc' "
	qsHc=qsHc&",(case AUTORIZACION.SOLO_CERTIFICADOS when 1 then 'Si' else 'No' end) as 'soloCert' " 
	qsHc=qsHc&",AUTORIZACION.ID_PROGRAMA,AUTORIZACION.ID_BLOQUE,AUTORIZACION.TIPO_DOC,CURRICULO.ID_CECO,AUTORIZACION.ID_AUTORIZACION,Factura=[dbo].[DET_FACTURAS_AUTORIZACION](AUTORIZACION.ID_AUTORIZACION), BANCOS.NOMBRE_BANCO, CUENTAS_CORRIENTES.NUMERO_CUENTA,AUTORIZACION.MONTO_TRANSFERENCIA, CONVERT(VARCHAR(10),AUTORIZACION.FECHA_TRANSFERENCIA, 105) as FECHA_TRANSFERENCIA, AUTORIZACION.RUT_DEPOSITANTE, AUTORIZACION.NOMBRE_DEPOSITANTE, AUTORIZACION.EMAIL_DEPOSITANTE, AUTORIZACION.RELACION_DEPOSITANTE,"
        qsHc=qsHc&"'Tipo Venta'=(case when AUTORIZACION.ID_TIPO_VENTA=1 then 'Venta Directa' when AUTORIZACION.ID_TIPO_VENTA=2 then 'Venta Broker' "
	qsHc=qsHc&" else 'Venta Directa' end)   from AUTORIZACION "
	qsHc=qsHc&" inner join PROGRAMA on PROGRAMA.ID_PROGRAMA=AUTORIZACION.ID_PROGRAMA "
	qsHc=qsHc&" left join BANCOS on BANCOS.ID_BANCO=AUTORIZACION.ID_BANCO "
	qsHc=qsHc&" left join CUENTAS_CORRIENTES on CUENTAS_CORRIENTES.ID_BANCO=AUTORIZACION.ID_BANCO AND CUENTAS_CORRIENTES.ID_CUENTA_CORRIENTE=AUTORIZACION.ID_CUENTA_CORRIENTE  "
	qsHc=qsHc&" inner join CURRICULO on CURRICULO.ID_MUTUAL=PROGRAMA.ID_MUTUAL "
	qsHc=qsHc&" inner join EMPRESAS on EMPRESAS.ID_EMPRESA=AUTORIZACION.ID_EMPRESA "
	qsHc=qsHc&" inner join bloque_programacion on bloque_programacion.id_bloque=AUTORIZACION.ID_BLOQUE "
	qsHc=qsHc&" inner join INSTRUCTOR_RELATOR on INSTRUCTOR_RELATOR.ID_INSTRUCTOR=bloque_programacion.id_relator "
	qsHc=qsHc&" where "&insSel&mesSel&anoSel
	qsHc=qsHc&" order by PROGRAMA.FECHA_INICIO_ asc "

	set rsHc =  conn.execute(qsHc)%>

    <table width="4000" border="1">
    <tr>
      <td width="100" align="center"><b><font size="2">Mes</font></b></td>
      <td width="50" align="center"><b><font size="2">Año</font></b></td>
      <td width="100" align="center"><b><font size="2">Rut de Empresa</font></b></td>
      <td width="350" align="center"><b><font size="2">Razón Social</font></b></td>
      <td width="80" align="center"><b><font size="2">Valor</font></b></td>
      <td width="200" align="center"><b><font size="2">Tipo de Documento</font></b></td>
      <td width="150" align="center"><b><font size="2">N° de Documento</font></b></td>      
      <td width="200" align="center"><b><font size="2">Con Franquicia</font></b></td>   
      <td width="100" align="center"><b><font size="2">C&oacute;digo de Curso</font></b></td> 
      <td width="80" align="center"><b><font size="2">Id Ceco</font></b></td> 
      <td width="550" align="center"><b><font size="2">Nombre del Curso</font></b></td> 
      <td width="250" align="center"><b><font size="2">Relator</font></b></td>              
      <td width="100" align="center"><b><font size="2">Fecha de Inicio</font></b></td>
      <td width="120" align="center"><b><font size="2">Fecha de T&eacute;rmino</font></b></td>
      <td width="80" align="center"><b><font size="2">N° Part.</font></b></td>
      <td width="100" align="center"><b><font size="2">Ver Documento</font></b></td>   
      <td width="150" align="center"><b><font size="2">Estado</font></b></td>      
      <td width="150" align="center"><b><font size="2">Libera Certificados</font></b></td>   
      <td width="150" align="center"><b><font size="2">ID Autorizaci&oacute;n</font></b></td>  
   <td width="100" align="center"><b><font size="2">Factura</font></b></td> 
      <td width="150" align="center"><b><font size="2">Tipo Venta</font></b></td>  
	  <td width="150" align="center"><b><font size="2">Banco</font></b></td>  
	  <td width="150" align="center"><b><font size="2">N. Cuenta</font></b></td>  
	  <td width="150" align="center"><b><font size="2">Monto Transferencia</font></b></td>  
	  <td width="150" align="center"><b><font size="2">Fecha Transferencia</font></b></td>   
	  <td width="150" align="center"><b><font size="2">Rut Depositante</font></b></td>  
	  <td width="150" align="center"><b><font size="2">Nombre Depositante</font></b></td>  
	  <td width="150" align="center"><b><font size="2">Email Depositante</font></b></td>  
	  <td width="150" align="center"><b><font size="2">Relaci&oacute;n Depositante</font></b></td> 
    </tr>
	<%
    while not rsHc.eof
    %>
    <tr>
      <td align="center"><b><font size="2"><%=rsHc("Mes")%></font></b></td>
      <td align="center"><b><font size="2"><%=rsHc("Año")%></font></b></td>     
      <td align="right"><b><font size="2"><%=rsHc("RUT")%></font></b></td>
      <td><b><font size="2"><%=rsHc("RAZÓN SOCIAL")%></font></b></td>
      <td><b><font size="2"><%=rsHc("VALOR_OC")%></font></b></td>
      <td><b><font size="2"><%=rsHc("Tipo Documento")%></font></b></td>
      <td align="right"><b><font size="2"><%=rsHc("ORDEN_COMPRA")%></font></b></td>  
      <td><b><font size="2"><%=rsHc("CF")%></font></b></td>      
      <td align="right"><b><font size="2"><%=rsHc("CODIGO")%></font></b></td> 
      <td align="center"><b><font size="2"><%=rsHc("id_ceco")%></font></b></td>  
      <td><b><font size="2"><%=rsHc("DESCRIPCION")%></font></b></td>
      <td><b><font size="2"><%=rsHc("instructor")%></font></b></td>
      <td align="center"><b><font size="2"><%=rsHc("FECHA_INICIO_")%></font></b></td>
      <td align="center"><b><font size="2"><%=rsHc("FECHA_TERMINO")%></font></b></td>
      <td align="center"><b><font size="2"><%=rsHc("total")%></font></b></td>
      <td align="center"><b><font size="2"><a href="<%=rsHc("DOCUMENTO_COMPROMISO")%>">Ver</a></font></b></td>  
      <td align="center"><b><font size="2"><%=rsHc("estInsc")%></font></b></td>
      <td align="center"><b><font size="2"><%=rsHc("soloCert")%></font></b></td>
      <td align="center"><b><font size="2"><%=rsHc("ID_AUTORIZACION")%></font></b></td>
        <td><b><font size="2"><%=rsHc("FACTURA")%></font></b></td>  
      <td align="center"><b><font size="2"><%=rsHc("Tipo Venta")%></font></b></td>
	  <td align="center"><b><font size="2"><%=rsHc("NOMBRE_BANCO")%></font></b></td>
	  <td align="center"><b><font size="2"><%=rsHc("NUMERO_CUENTA")%></font></b></td>
	  <td align="center"><b><font size="2"><%=rsHc("MONTO_TRANSFERENCIA")%></font></b></td>
	  <td align="center"><b><font size="2"><%=rsHc("FECHA_TRANSFERENCIA")%></font></b></td>
	  <td align="center"><b><font size="2"><%=rsHc("RUT_DEPOSITANTE")%></font></b></td>
	  <td align="center"><b><font size="2"><%=rsHc("NOMBRE_DEPOSITANTE")%></font></b></td>
	  <td align="center"><b><font size="2"><%=rsHc("EMAIL_DEPOSITANTE")%></font></b></td>
	  <td align="center"><b><font size="2"><%=rsHc("RELACION_DEPOSITANTE")%></font></b></td>
    </tr>	
    <%
        rsHc.MoveNext
    wend
    %>
    </table>
