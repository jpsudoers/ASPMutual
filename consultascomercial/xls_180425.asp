<!--#include file="../conexion.asp"-->
<%
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
	
qsMan="select PROGRAMA.ID_PROGRAMA ,CURRICULO.CODIGO ,PROGRAMA.FECHA_INICIO_, EM.rut as rut_empresa, dbo.MayMinTexto (CURRICULO.NOMBRE_CURSO) as NOMBRE_CURSO ,CONVERT(VARCHAR(10),PROGRAMA.FECHA_INICIO_, 105) as FECHA_INICIO ,CONVERT(VARCHAR(10),PROGRAMA.FECHA_TERMINO, 105) as FECHA_TERMINO ,PROGRAMA.CUPOS, (select COUNT(*) from HISTORICO_CURSOS where HISTORICO_CURSOS.ID_PROGRAMA=PROGRAMA.ID_PROGRAMA) as inscritos, (select COUNT(*) from HISTORICO_CURSOS where HISTORICO_CURSOS.ID_PROGRAMA=PROGRAMA.ID_PROGRAMA and HISTORICO_CURSOS.ASISTENCIA>0 and HISTORICO_CURSOS.ESTADO=0) as Asistentes, (select COUNT(*) from HISTORICO_CURSOS where HISTORICO_CURSOS.ID_PROGRAMA=PROGRAMA.ID_PROGRAMA and HISTORICO_CURSOS.EVALUACION='Aprobado' and HISTORICO_CURSOS.ESTADO=0) as Aprobados, (select top 1 (CASE WHEN BQ.id_sede<>27 THEN (SELECT ISNULL(Z.nombre,'') FROM sedes s INNER JOIN ZONA Z ON Z.id_zona = S.ID_ZONA WHERE s.ID_SEDE=bq.id_sede) ELSE (SELECT ISNULL(ZN.nombre,'') FROM REGIONES R INNER JOIN ZONA ZN ON ZN.id_zona = R.ID_ZONA WHERE R.ID_REGION=BQ.id_region) END) from bloque_programacion bq where bq.id_programa=PROGRAMA.ID_PROGRAMA) AS ZONAS, (select COUNT(*) from HISTORICO_CURSOS where HISTORICO_CURSOS.ID_PROGRAMA=PROGRAMA.ID_PROGRAMA and HISTORICO_CURSOS.EVALUACION='Reprobado' and HISTORICO_CURSOS.ESTADO=0) as Reprobados, PROGRAMA.TIPO AS TIPO ,(case when PROGRAMA.DIR_EJEC is not null then PROGRAMA.DIR_EJEC else '' end) as SEDE,(select top 1 case when bq.id_sede<>27 then s.NOMBRE+', '+s.DIRECCION+', '+s.CIUDAD when bq.id_sede=27 then bq.nom_sede end from bloque_programacion bq  inner join sedes s on s.ID_SEDE=bq.id_sede where bq.id_programa=PROGRAMA.ID_PROGRAMA) as sede_curso ,(select top 1 case when bq.id_sede<>27 then ISNULL((select cm.NOMBRE from comuna cm where cm.ID_COMUNA=s.ID_COMUNA),'') when bq.id_sede=27 then ISNULL((select cm.NOMBRE from comuna cm where cm.ID_COMUNA=bq.id_comuna),'') end from bloque_programacion bq inner join sedes s on s.ID_SEDE=bq.id_sede where bq.id_programa=PROGRAMA.ID_PROGRAMA) as 'Comuna' ,(select top 1 case when bq.id_sede<>27 then ISNULL((select r.NOMBRE from regiones r where r.ID_REGION=s.ID_REGION),'') when bq.id_sede=27 then ISNULL((select r.NOMBRE from regiones r where r.ID_REGION=bq.id_region),'') end from bloque_programacion bq inner join sedes s on s.ID_SEDE=bq.id_sede where bq.id_programa=PROGRAMA.ID_PROGRAMA) as 'Region' ,(select top 1 case when bq.id_sede<>27 then ISNULL((select r.id_region from regiones r where r.ID_REGION=s.ID_REGION),'') when bq.id_sede=27 then ISNULL((select r.id_region from regiones r where r.ID_REGION=bq.id_region),'') end from bloque_programacion bq inner join sedes s on s.ID_SEDE=bq.id_sede where bq.id_programa=PROGRAMA.ID_PROGRAMA) as 'idRegion' ,(select top 1 s.ID_COD_INT_MUTUAL from bloque_programacion bq inner join sedes s on s.ID_SEDE=bq.id_sede where bq.id_programa=PROGRAMA.ID_PROGRAMA) as id_sede_curso ,HI=(select hb.HORARIO from HORARIO_BLOQUES hb where hb.ID_HORARIO=PROGRAMA.BMI) ,HF=(select hb.HORARIO from HORARIO_BLOQUES hb where hb.ID_HORARIO=PROGRAMA.BTF) ,IR.RUT as Rut_instructor ,IR.NOMBRES+' '+Ir.A_PATERNO+' '+Ir.A_MATERNO AS instructor ,CURRICULO.ID_PROYECTO ,(CASE CURRICULO.ID_PROYECTO WHEN 1 THEN 'Abierta' WHEN 2 THEN 'Cerrada' WHEN 3 THEN 'Proyecto Almagro' WHEN 4 THEN 'Spot Mutual' WHEN 5 THEN 'Cerrado-Collahuasi' ELSE '' end) as 'PROYECTO' ,(select top 1 dbo.Estado_evaluacion(bq.estado_eva_cdn) from bloque_programacion bq where bq.id_programa=PROGRAMA.ID_PROGRAMA) as 'EC' ,  usr=ISNULL((select u.NOMBRES+' '+u.A_PATERNO+' '+u.A_MATERNO from usuarios u where u.ID_USUARIO=PROGRAMA.USUARIO_CREA),'') , PROGRAMA.ESTADO ,CURRICULO.PGP ,VW.*, EM.ACT_ECONOMICA, AE.ACT_ECO_DESCRIPCION, AE.GRUPO_TEXTO, ISNULL(TI.TIPO_INDUCCION_DESC,'') TIPO_INDUCCION_DESC from PROGRAMA inner join CURRICULO on CURRICULO.ID_MUTUAL=PROGRAMA.ID_MUTUAL LEFT JOIN bloque_programacion BQ ON BQ.id_programa = PROGRAMA.ID_PROGRAMA LEFT JOIN INSTRUCTOR_RELATOR IR ON ir.ID_INSTRUCTOR=bq.id_relator LEFT JOIN VW_PIVOT_GASTO_PROGRAMA VW ON VW.ID_BLOQUE = BQ.id_bloque left join empresas EM on EM.ID_EMPRESA = programa.ID_EMPRESA LEFT JOIN ACTIVIDAD_ECONOMICA_EMPRESA AE ON AE.ACT_ECO_CODIGO=EM.ACT_ECONOMICA LEFT JOIN TIPO_INDUCCION TI ON TI.ID_TIPO_INDUCCION=CURRICULO.ID_TIPO_INDUCCION where PROGRAMA.ESTADO=1 AND PROGRAMA.ID_CLIENTE in ("&request("c")&")  "&proSel&mesSel&anoSel& "order by programa.FECHA_INICIO_"
	  'and programa.id_programa < 22232" where PROGRAMA.FECHA_INICIO_>=CONVERT(date,'"&request("f_ini")&"') "&_
	  '" and PROGRAMA.FECHA_INICIO_<DATEADD(DAY, 1, CONVERT(date,'"&request("f_fin")&"', 105))"
	  
	'Response.Write(qsMan)
	'Response.End()

	set rsMan =  conn.execute(qsMan)
	

	%>
    
    

    <table width="800" border="0">
        <tr>
          <td width="60" align="center" style="border-width: 1px;border: solid; border-color:#000;"><b><font size="2">ID</font></b></td>
		  <td width="150" align="center" style="border-width: 1px;border: solid; border-color:#000;"><b><font size="2">CODIGO CURSO</font></b></td>   
          <td width="190" align="center" style="border-width: 1px;border: solid; border-color:#000;"><b><font size="2">RUT_EMPRESA</font></b></td>            
          <td width="850" align="center" style="border-width: 1px;border: solid; border-color:#000;"><b><font size="2">NOMBRE CURSO</font></b></td>
		  <td width="150" align="center" style="border-width: 1px;border: solid; border-color:#000;"><b><font size="2">FECHA INICIO</font></b></td>           		
          <td width="150" align="center" style="border-width: 1px;border: solid; border-color:#000;"><b><font size="2">FECHA TERMINO</font></b></td>
		  <td width="130" align="center" style="border-width: 1px;border: solid; border-color:#000;"><b><font size="2">CUPOS</font></b></td>
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
 		  <td width="350" align="center" style="border-width: 1px;border: solid; border-color:#000;"><b><font size="2">MATERIAL</font></b></td>  
 		  <td width="350" align="center" style="border-width: 1px;border: solid; border-color:#000;"><b><font size="2">FANTOMA</font></b></td>  
 		  <td width="350" align="center" style="border-width: 1px;border: solid; border-color:#000;"><b><font size="2">SIMULADOR</font></b></td>
 		  <td width="350" align="center" style="border-width: 1px;border: solid; border-color:#000;"><b><font size="2">PROYECTOR</font></b></td>
 		  <td width="350" align="center" style="border-width: 1px;border: solid; border-color:#000;"><b><font size="2">ALMUERZO</font></b></td>
 		  <td width="350" align="center" style="border-width: 1px;border: solid; border-color:#000;"><b><font size="2">OTROS</font></b></td>          
          <td width="350" align="center" style="border-width: 1px;border: solid; border-color:#000;"><b><font size="2">ACT. ECONOMICA</font></b></td>
          <td width="350" align="center" style="border-width: 1px;border: solid; border-color:#000;"><b><font size="2">ACT. ECO. DESCRIPCION</font></b></td>
          <td width="350" align="center" style="border-width: 1px;border: solid; border-color:#000;"><b><font size="2">ACT. ECO. GRUPO</font></b></td>                                                                    
          <td width="350" align="center" style="border-width: 1px;border: solid; border-color:#000;"><b><font size="2">TIPO CURSO</font
><!--FIN costos del curso-->          
                                      
        </tr>
        <%
		
        while not rsMan.eof
        %>
        <tr>
          <td style="border-width: 1px;border: solid; border-color:#000;"><font size="2"><%=rsMan("ID_PROGRAMA")%></font></td>
          <td style="border-width: 1px;border: solid; border-color:#000;"><font size="2"><%=rsMan("CODIGO")%></font></td>  
           <td style="border-width: 1px;border: solid; border-color:#000;"><font size="2"><%=rsMan("RUT_EMPRESA")%></font></td>                 
          <td style="border-width: 1px;border: solid; border-color:#000;"><font size="2"><%=rsMan("NOMBRE_CURSO")%></font></td>
          <td align="center" style="border-width: 1px;border: solid; border-color:#000;"><font size="2"><%=rsMan("FECHA_INICIO")%></font></td>  
          <td align="center" style="border-width: 1px;border: solid; border-color:#000;"><font size="2"><%=rsMan("FECHA_TERMINO")%></font></td>
	      <td style="border-width: 1px;border: solid; border-color:#000;"><font size="2"><%=rsMan("CUPOS")%></font></td>	
          <td style="border-width: 1px;border: solid; border-color:#000;"><font size="2"><%=rsMan("inscritos")%></font></td>   
            
          <td style="border-width: 1px;border: solid; border-color:#000;"><font size="2"><%=rsMan("Asistentes")%></font></td>
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
 		  <td style="border-width: 1px;border: solid; border-color:#000;"><font size="2"><%=rsMan("MATERIAL")%></font></b></td>  
 		  <td style="border-width: 1px;border: solid; border-color:#000;"><font size="2"><%=rsMan("FANTOMA")%></font></b></td>  
 		  <td style="border-width: 1px;border: solid; border-color:#000;"><font size="2"><%=rsMan("SIMULADOR")%></font></b></td>
 		  <td style="border-width: 1px;border: solid; border-color:#000;"><font size="2"><%=rsMan("PROYECTOR")%></font></b></td>
 		  <td style="border-width: 1px;border: solid; border-color:#000;"><font size="2"><%=rsMan("ALMUERZO")%></font></b></td>
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
