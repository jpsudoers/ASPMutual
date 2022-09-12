<!--#include file="../conexion.asp"-->
<%

Response.CodePage = 65001
Response.CharSet = "utf-8"

	dim mes(12) 
	mes(0)="Enero"
	mes(1)="Febrero"
	mes(2)="Marzo"
	mes(3)="Abril"
	mes(4)="Mayo"
	mes(5)="Junio"
	mes(6)="Julio"
	mes(7)="Agosto"
	mes(8)="Septiembre"
	mes(9)="Octubre"
	mes(10)="Noviembre"
	mes(11)="Diciembre"

	fecha = year(now)&right("0"&month(now()),2)&right("0"&day(now()),2)
	fecha = fecha&right("0"&hour(now()),2)&right("0"&minute(now()),2)&right("0"&second(now()),2)
	Response.ContentType ="application/vnd.ms-excel"
	Response.AddHeader "content-disposition", "inline; filename=Informe_Comercial_"&fecha&".xls"
	
	mesSel=""
	if(request("mes")<>"0")then
		mesSel=" and Month(PROGRAMA.FECHA_INICIO_)="&request("mes")
	end if

	anoSel=""
	if(request("ano")<>"0")then
		anoSel=" and Year(PROGRAMA.FECHA_INICIO_)="&request("ano")
	end if	
	
	proSel=""
	if(request("p")<>"0")then
		proSel=" and CURRICULO.ID_PROYECTO="&request("p")
	end if		
	
qsMan="select PROGRAMA.ID_PROGRAMA ,CURRICULO.CODIGO ,PROGRAMA.FECHA_INICIO_, "&_
		"EM.rut as rut_empresa,  "&_
		"dbo.MayMinTexto (CURRICULO.NOMBRE_CURSO) as NOMBRE_CURSO , "&_
		"CONVERT(VARCHAR(10),PROGRAMA.FECHA_INICIO_, 105) as FECHA_INICIO , "&_
		"CONVERT(VARCHAR(10),PROGRAMA.FECHA_TERMINO, 105) as FECHA_TERMINO , "&_
		"PROGRAMA.CUPOS,  "&_
		"(select COUNT(*) from HISTORICO_CURSOS  "&_
		" where HISTORICO_CURSOS.ID_PROGRAMA=PROGRAMA.ID_PROGRAMA AND HISTORICO_CURSOS.ID_BLOQUE=BQ.id_bloque) as inscritos,  "&_
		"(select COUNT(*) from HISTORICO_CURSOS  "&_
		" where HISTORICO_CURSOS.ID_PROGRAMA=PROGRAMA.ID_PROGRAMA  "&_
		" and HISTORICO_CURSOS.ASISTENCIA>0 and HISTORICO_CURSOS.ESTADO=0  AND HISTORICO_CURSOS.ID_BLOQUE=BQ.id_bloque) as Asistentes, "&_
		" (select COUNT(*) from HISTORICO_CURSOS  "&_
		" where HISTORICO_CURSOS.ID_PROGRAMA=PROGRAMA.ID_PROGRAMA  "&_
		" and HISTORICO_CURSOS.EVALUACION='Aprobado'  "&_
		" and HISTORICO_CURSOS.ESTADO=0 AND HISTORICO_CURSOS.ID_BLOQUE=BQ.id_bloque) as Aprobados,  "&_
		" (CASE WHEN BQ.id_sede<>27 THEN  "&_
		" (SELECT ISNULL(Z.nombre,'') FROM ZONA Z  "&_
		" WHERE Z.id_zona = S.ID_ZONA)  "&_
		" ELSE (SELECT ISNULL(ZN.nombre,'') FROM REGIONES R  "&_
		" INNER JOIN ZONA ZN ON ZN.id_zona = R.ID_ZONA  "&_
		" WHERE R.ID_REGION=BQ.id_region) END) AS ZONAS,  "&_
		" (select COUNT(*) from HISTORICO_CURSOS  "&_
		" where HISTORICO_CURSOS.ID_PROGRAMA=PROGRAMA.ID_PROGRAMA  "&_
		" and HISTORICO_CURSOS.EVALUACION='Reprobado'  "&_
		" and HISTORICO_CURSOS.ESTADO=0 AND HISTORICO_CURSOS.ID_BLOQUE=BQ.id_bloque) as Reprobados,  "&_
		" PROGRAMA.TIPO AS TIPO, "&_
		" (case when PROGRAMA.DIR_EJEC is not null then  "&_
		" PROGRAMA.DIR_EJEC else '' end) as SEDE, "&_
		" (case when bq.id_sede<>27 then  "&_
		" s.NOMBRE+', '+s.DIRECCION+', '+s.CIUDAD  "&_
		" when bq.id_sede=27 then bq.nom_sede end) as sede_curso , "&_
		" (case when bq.id_sede<>27 then  "&_
		" ISNULL((select cm.NOMBRE from comuna cm  "&_
		" where cm.ID_COMUNA=s.ID_COMUNA),'')  "&_
		" when bq.id_sede=27 then  "&_
		" ISNULL((select cm.NOMBRE from comuna cm  "&_
		" where cm.ID_COMUNA=bq.id_comuna),'') end) as 'Comuna' , "&_
		" (case when bq.id_sede<>27 then  "&_
		" ISNULL((select r.NOMBRE from regiones r  "&_
		" where r.ID_REGION=s.ID_REGION),'') when bq.id_sede=27 then  "&_
		" ISNULL((select r.NOMBRE from regiones r  "&_
		" where r.ID_REGION=bq.id_region),'') end) as 'Region' , "&_
		" ( case when bq.id_sede<>27 then  "&_
		" ISNULL((select r.id_region from regiones r  "&_
		" where r.ID_REGION=s.ID_REGION),'') when bq.id_sede=27 then  "&_
		" ISNULL((select r.id_region from regiones r  "&_
		" where r.ID_REGION=bq.id_region),'') end) as 'idRegion' , "&_
		" (s.ID_COD_INT_MUTUAL) as id_sede_curso , "&_
		" HI=(select hb.HORARIO from HORARIO_BLOQUES hb  "&_
		" where hb.ID_HORARIO=PROGRAMA.BMI) , "&_
		" HF=(select hb.HORARIO from HORARIO_BLOQUES hb  "&_
		" where hb.ID_HORARIO=PROGRAMA.BTF) ,IR.RUT as Rut_instructor , "&_
		" IR.NOMBRES+' '+Ir.A_PATERNO+' '+Ir.A_MATERNO AS instructor , "&_
		" CURRICULO.ID_PROYECTO, "&_
		" (CASE CURRICULO.ID_PROYECTO "&_
		" WHEN 1 THEN 'Abierta' WHEN 2 THEN 'Cerrada' "&_
		" WHEN 3 THEN 'Proyecto Almagro' WHEN 4 THEN 'Spot Mutual' "&_
		" WHEN 5 THEN 'Cerrado-Collahuasi' ELSE '' end) as 'PROYECTO', "&_
		" (dbo.Estado_evaluacion(isnull(bq.estado_eva_cdn,0))) as 'EC', "&_
		" usr=ISNULL((select u.NOMBRES+' '+u.A_PATERNO+' '+u.A_MATERNO "&_
		" from usuarios u where u.ID_USUARIO=PROGRAMA.USUARIO_CREA),''), "&_
		" PROGRAMA.ESTADO ,CURRICULO.PGP ,VW.*, EM.ACT_ECONOMICA, "&_
		" AE.ACT_ECO_DESCRIPCION, AE.GRUPO_TEXTO,  "&_
		" ISNULL(TI.TIPO_INDUCCION_DESC,'') TIPO_INDUCCION_DESC "&_
		" , PROGRAMA.ASIS_RELATOR as 'ASISRELATOR' " &_
		" , CASE ISNULL(PARSENAME(REPLACE(PROGRAMA.DISPOSITIVO_RELATOR, ';', '.'), 2), '') WHEN ' Desktop' THEN 'Computador' WHEN ' Smartphone' THEN 'Celular' ELSE (CASE WHEN PROGRAMA.ASIS_RELATOR >= 0 THEN 'Computador' ELSE '' END) END as 'DISPRELATOR',/*CURRICULO.ID_CECO*/" &_
		"ID_CECO=(case when PROGRAMA.ID_CLIENTE=13 and (PROGRAMA.ID_MUTUAL=4794 or "&_
		"PROGRAMA.ID_MUTUAL=4795) then " &_ 
		"(case when bq.id_sede<>27 then [dbo].[AsignaCCAha](s.ID_REGION,CURRICULO.ID_MUTUAL) "&_
		" when bq.id_sede=27 then [dbo].[AsignaCCAha](bq.id_region,CURRICULO.ID_MUTUAL) end) "&_
		" else CURRICULO.ID_CECO end) "&_
		" from PROGRAMA  "&_
		" inner join CURRICULO on CURRICULO.ID_MUTUAL=PROGRAMA.ID_MUTUAL  "&_
		" LEFT JOIN bloque_programacion BQ ON BQ.id_programa = PROGRAMA.ID_PROGRAMA  "&_
		" LEFT JOIN INSTRUCTOR_RELATOR IR ON ir.ID_INSTRUCTOR=bq.id_relator  "&_
		" LEFT join sedes s on s.ID_SEDE=bq.id_sede  "&_
		" LEFT JOIN VW_PIVOT_GASTO_PROGRAMA VW ON VW.ID_BLOQUE = BQ.id_bloque  "&_
		" left join empresas EM on EM.ID_EMPRESA = programa.ID_EMPRESA  "&_
		" LEFT JOIN ACTIVIDAD_ECONOMICA_EMPRESA AE ON AE.ACT_ECO_CODIGO=EM.ACT_ECONOMICA  "&_
		" LEFT JOIN TIPO_INDUCCION TI ON TI.ID_TIPO_INDUCCION=CURRICULO.ID_TIPO_INDUCCION  "&_
		" where PROGRAMA.ESTADO=1 AND PROGRAMA.ID_CLIENTE in ("&request("c")&") "&proSel&mesSel&anoSel&_
		" order by programa.FECHA_INICIO_"

	set rsMan =  conn.execute(qsMan)	

	%> 
    
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />    
    <table width="800" border="0">
    <tr>
    <td width="60" align="center" style="border-width: 1px;border: solid; border-color:#000;"><b><font size="2">ID</font></b></td>
	<td width="150" align="center" style="border-width: 1px;border: solid; border-color:#000;"><b><font size="2">CENTRO COSTO</font></b></td>      
	<td width="150" align="center" style="border-width: 1px;border: solid; border-color:#000;"><b><font size="2">CODIGO CURSO</font></b></td>   
    <td width="190" align="center" style="border-width: 1px;border: solid; border-color:#000;"><b><font size="2">RUT_EMPRESA</font></b></td>            
    <td width="850" align="center" style="border-width: 1px;border: solid; border-color:#000;"><b><font size="2">NOMBRE CURSO</font></b></td>
		  <td width="150" align="center" style="border-width: 1px;border: solid; border-color:#000;"><b><font size="2">FECHA INICIO</font></b></td>           		
          <td width="150" align="center" style="border-width: 1px;border: solid; border-color:#000;"><b><font size="2">FECHA TERMINO</font></b></td>
		  <td width="130" align="center" style="border-width: 1px;border: solid; border-color:#000;"><b><font size="2">CUPOS</font></b></td>
         <td width="100" align="center" style="border-width: 1px;border: solid; border-color:#000;"><b><font size="2">ASISTENCIA RELATOR</font></b></td>
          <td width="100" align="center" style="border-width: 1px;border: solid; border-color:#000;"><b><font size="2">DISPOSITIVO RELATOR</font></b></td>
		 <td width="130" align="center" style="border-width: 1px;border: solid; border-color:#000;"><b><font size="2">INSCRITOS</font></b></td>                     
          <td width="130" align="center" style="border-width: 1px;border: solid; border-color:#000;"><b><font size="2">ASISTENTES</font></b></td>
		  <td width="130" align="center" style="border-width: 1px;border: solid; border-color:#000;"><b><font size="2">APROBADOS</font></b></td>           
          <td width="130" align="center" style="border-width: 1px;border: solid; border-color:#000;"><b><font size="2">REPROBADOS</font></b></td>
          <td width="700" align="center" style="border-width: 1px;border: solid; border-color:#000;"><b><font size="2">ZONA</font></b></td> 
		  <td width="700" align="center" style="border-width: 1px;border: solid; border-color:#000;"><b><font size="2">SEDE</font></b></td>      
		 <td width="700" align="center" style="border-width: 1px;border: solid; border-color:#000;"><b><font size="2">SEDE CURSO</font></b></td>       
		 <td width="400" align="center" style="border-width: 1px;border: solid; border-color:#000;"><b><font size="2">REGION</font></b></td>   
		 <td width="400" align="center" style="border-width: 1px;border: solid; border-color:#000;"><b><font size="2">COMUNA</font></b></td>                                    
          <td width="100" align="center" style="border-width: 1px;border: solid; border-color:#000;"><b><font size="2">AGENCIA</font></b></td>
		  <td width="80" align="center" style="border-width: 1px;border: solid; border-color:#000;"><b><font size="2">HI</font></b></td>           
          <td width="80" align="center" style="border-width: 1px;border: solid; border-color:#000;"><b><font size="2">HF</font></b></td>
		  <td width="150" align="center" style="border-width: 1px;border: solid; border-color:#000;"><b><font size="2">RUT INSTRUCTOR</font></b></td> 
          <td width="400" align="center" style="border-width: 1px;border: solid; border-color:#000;"><b><font size="2">INSTRUCTOR</font></b></td>
		  <td width="200" align="center" style="border-width: 1px;border: solid; border-color:#000;"><b><font size="2">PROYECTO</font></b></td> 
		  <td width="80" align="center" style="border-width: 1px;border: solid; border-color:#000;"><b><font size="2">PGP</font></b></td>   
		  <td width="350" align="center" style="border-width: 1px;border: solid; border-color:#000;"><b><font size="2">ESTADO CURSO</font></b></td>     
		  <td width="350" align="center" style="border-width: 1px;border: solid; border-color:#000;"><b><font size="2">COORDINADO POR</font></b></td>  
<!--INICIO costos del curso-->         
 		  <td width="350" align="center" style="border-width: 1px;border: solid; border-color:#000;"><b><font size="2">ARRIENDO DE SALAS</font></b></td>          
 		  <td width="350" align="center" style="border-width: 1px;border: solid; border-color:#000;"><b><font size="2">RELATORIA</font></b></td>  
 		  <td width="350" align="center" style="border-width: 1px;border: solid; border-color:#000;"><b><font size="2">COFFEE</font></b></td>  
 		  <td width="350" align="center" style="border-width: 1px;border: solid; border-color:#000;"><b><font size="2">TRASLADO</font></b></td>  
 		  <!--<td width="350" align="center" style="border-width: 1px;border: solid; border-color:#000;"><b><font size="2">MATERIAL</font></b></td>  
 		  <td width="350" align="center" style="border-width: 1px;border: solid; border-color:#000;"><b><font size="2">FANTOMA</font></b></td>  
 		  <td width="350" align="center" style="border-width: 1px;border: solid; border-color:#000;"><b><font size="2">SIMULADOR</font></b></td>
 		  <td width="350" align="center" style="border-width: 1px;border: solid; border-color:#000;"><b><font size="2">PROYECTOR</font></b></td>
 		  <td width="350" align="center" style="border-width: 1px;border: solid; border-color:#000;"><b><font size="2">ALMUERZO</font></b></td>-->
 		  <td width="350" align="center" style="border-width: 1px;border: solid; border-color:#000;"><b><font size="2">OTROS</font></b></td>          
          <td width="350" align="center" style="border-width: 1px;border: solid; border-color:#000;"><b><font size="2">ACT. ECONOMICA</font></b></td>
          <td width="350" align="center" style="border-width: 1px;border: solid; border-color:#000;"><b><font size="2">ACT. ECO. DESCRIPCION</font></b></td>
          <td width="350" align="center" style="border-width: 1px;border: solid; border-color:#000;"><b><font size="2">ACT. ECO. GRUPO</font></b></td>                                                                    
          <td width="350" align="center" style="border-width: 1px;border: solid; border-color:#000;"><b><font size="2">TIPO CURSO</font>          
          <!--FIN costos del curso-->          
                                      
        </tr>
        <%
		
        while not rsMan.eof
        %>
        <tr>
          <td style="border-width: 1px;border: solid; border-color:#000;"><font size="2"><%=rsMan("ID_PROGRAMA")%></font></td>
          <td style="border-width: 1px;border: solid; border-color:#000;"><font size="2"><%=rsMan("ID_CECO")%></font></td>            
          <td style="border-width: 1px;border: solid; border-color:#000;"><font size="2"><%=rsMan("CODIGO")%></font></td>  
          <td style="border-width: 1px;border: solid; border-color:#000;"><font size="2"><%=rsMan("RUT_EMPRESA")%></font></td>
          <td style="border-width: 1px;border: solid; border-color:#000;"><font size="2"><%=rsMan("NOMBRE_CURSO")%></font></td>
          <td align="center" style="border-width: 1px;border: solid; border-color:#000;"><font size="2"><%=rsMan("FECHA_INICIO")%></font></td>  
          <td align="center" style="border-width: 1px;border: solid; border-color:#000;"><font size="2"><%=rsMan("FECHA_TERMINO")%></font></td>
	      <td style="border-width: 1px;border: solid; border-color:#000;"><font size="2"><%=rsMan("CUPOS")%></font></td>
          
          <td style="border-width: 1px;border: solid; border-color:#000;"><font size="2"><%=rsMan("ASISRELATOR")%></font></b></td>
		  <td style="border-width: 1px;border: solid; border-color:#000;"><font size="2"><%=rsMan("DISPRELATOR")%></font></b></td>
          
          <td style="border-width: 1px;border: solid; border-color:#000;"><font size="2"><%=rsMan("inscritos")%></font></td>          <td style="border-width: 1px;border: solid; border-color:#000;"><font size="2"><%=rsMan("Asistentes")%></font></td>
          <td style="border-width: 1px;border: solid; border-color:#000;"><font size="2"><%=rsMan("Aprobados")%></font></td>  
          <td style="border-width: 1px;border: solid; border-color:#000;"><font size="2"><%=rsMan("Reprobados")%></font></td>
	      <td style="border-width: 1px;border: solid; border-color:#000;"><font size="2"><%=rsMan("ZONAS")%></font></td>	
	      <td style="border-width: 1px;border: solid; border-color:#000;"><font size="2"><%=rsMan("SEDE")%></font></td>	          
          <td style="border-width: 1px;border: solid; border-color:#000;"><font size="2"><%=rsMan("sede_curso")%></font></td>  
          <td style="border-width: 1px;border: solid; border-color:#000;"><font size="2"><%=rsMan("Region")%></font></td>   
          <td style="border-width: 1px;border: solid; border-color:#000;"><font size="2"><%=rsMan("Comuna")%></font></td>
          <td style="border-width: 1px;border: solid; border-color:#000;"><font size="2"><%=rsMan("id_sede_curso")%></font></td>
          <td style="border-width: 1px;border: solid; border-color:#000;"><font size="2"><%=rsMan("HI")%></font></td>  
          <td style="border-width: 1px;border: solid; border-color:#000;"><font size="2"><%=rsMan("HF")%></font></td>
	  	  <td style="border-width: 1px;border: solid; border-color:#000;"><font size="2"><%=" "&rsMan("Rut_instructor")%></font></td>
	      <td style="border-width: 1px;border: solid; border-color:#000;"><font size="2"><%=rsMan("instructor")%></font></td>	
	      <td style="border-width: 1px;border: solid; border-color:#000;"><font size="2"><%=rsMan("PROYECTO")%></font></td>	
	      <td style="border-width: 1px;border: solid; border-color:#000;"><font size="2"><%=rsMan("PGP")%></font></td>
          <td style="border-width: 1px;border: solid; border-color:#000;"><font size="2"><%if((rsMan("idRegion")=1 or rsMan("idRegion")=2 or rsMan("idRegion")=3 or rsMan("idRegion")=4 or rsMan("idRegion")=15) and rsMan("PROYECTO")="Abierta")then%>Ejecutado por Mutual<%else%><%=rsMan("EC")%><%end if%></font></td>     	
          <td style="border-width: 1px;border: solid; border-color:#000;"><font size="2"><%=rsMan("usr")%></font></td>
          
  <!--INICIO costos del curso-->         
 		  <td style="border-width: 1px;border: solid; border-color:#000;"><font size="2"><%=rsMan("ARRIENDO_sala")%></font></b></td>          
 		  <td style="border-width: 1px;border: solid; border-color:#000;"><font size="2"><%=rsMan("RELATORIA")%></font></b></td>  
 		  <td style="border-width: 1px;border: solid; border-color:#000;"><font size="2"><%=rsMan("COFFEE")%></font></b></td>  
 		  <td style="border-width: 1px;border: solid; border-color:#000;"><font size="2"><%=rsMan("TRASLADO")%></font></b></td>  
 		  <!--<td style="border-width: 1px;border: solid; border-color:#000;"><font size="2"><%=rsMan("MATERIAL")%></font></b></td>  
 		  <td style="border-width: 1px;border: solid; border-color:#000;"><font size="2"><%=rsMan("FANTOMA")%></font></b></td>  
 		  <td style="border-width: 1px;border: solid; border-color:#000;"><font size="2"><%=rsMan("SIMULADOR")%></font></b></td>
 		  <td style="border-width: 1px;border: solid; border-color:#000;"><font size="2"><%=rsMan("PROYECTOR")%></font></b></td>
 		  <td style="border-width: 1px;border: solid; border-color:#000;"><font size="2"><%=rsMan("ALMUERZO")%></font></b></td>-->
 		  <td style="border-width: 1px;border: solid; border-color:#000;"><font size="2"><%=rsMan("OTROS")%></font></b></td>
          
<td style="border-width: 1px;border: solid; border-color:#000;"><font size="2"><%=rsMan("ACT_ECONOMICA")%></font></b></td>
<td style="border-width: 1px;border: solid; border-color:#000;"><font size="2"><%=rsMan("ACT_ECO_DESCRIPCION")%></font></b></td>
<td style="border-width: 1px;border: solid; border-color:#000;"><font size="2"><%=rsMan("GRUPO_TEXTO")%></font></b></td>
<td style="border-width: 1px;border: solid; border-color:#000;"><font size="2"><%=rsMan("TIPO_INDUCCION_DESC")%></font></b></td>
		  
<!--FIN costos del curso-->        
          	
        </tr>	
        <%
            rsMan.MoveNext
        wend
        %>
    </table>
