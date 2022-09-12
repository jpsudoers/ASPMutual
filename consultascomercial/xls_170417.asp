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
	
qsMan="select PROGRAMA.ID_PROGRAMA,CURRICULO.CODIGO,PROGRAMA.FECHA_INICIO_,"&_
	  " dbo.MayMinTexto (CURRICULO.NOMBRE_CURSO) as NOMBRE_CURSO,"&_
	  " CONVERT(VARCHAR(10),PROGRAMA.FECHA_INICIO_, 105) as FECHA_INICIO,"&_
	  " CONVERT(VARCHAR(10),PROGRAMA.FECHA_TERMINO, 105) as FECHA_TERMINO,PROGRAMA.CUPOS,"&_
	  " (select COUNT(*) from HISTORICO_CURSOS where HISTORICO_CURSOS.ID_PROGRAMA=PROGRAMA.ID_PROGRAMA) as inscritos,"&_ 
	  " (select COUNT(*) from HISTORICO_CURSOS where HISTORICO_CURSOS.ID_PROGRAMA=PROGRAMA.ID_PROGRAMA "&_
	  " and HISTORICO_CURSOS.ASISTENCIA>0 and HISTORICO_CURSOS.ESTADO=0) as Asistentes, "&_
	  " (select COUNT(*) from HISTORICO_CURSOS where HISTORICO_CURSOS.ID_PROGRAMA=PROGRAMA.ID_PROGRAMA "&_
	  " and HISTORICO_CURSOS.EVALUACION='Aprobado' and HISTORICO_CURSOS.ESTADO=0) as Aprobados, "&_
	  " (select COUNT(*) from HISTORICO_CURSOS where HISTORICO_CURSOS.ID_PROGRAMA=PROGRAMA.ID_PROGRAMA "&_
	  " and HISTORICO_CURSOS.EVALUACION='Reprobado' and HISTORICO_CURSOS.ESTADO=0) as Reprobados, "&_
	  " PROGRAMA.TIPO, (case when PROGRAMA.DIR_EJEC is not null then  "&_
	  " PROGRAMA.DIR_EJEC else '' end) as SEDE, "&_
	  " (select case when bq.id_sede<>27 then s.NOMBRE+', '+s.DIRECCION+', '+s.CIUDAD  "&_
	  " when bq.id_sede=27 then bq.nom_sede end from bloque_programacion bq  "&_
	  " inner join sedes s on s.ID_SEDE=bq.id_sede "&_
	  " where bq.id_programa=PROGRAMA.ID_PROGRAMA) as sede_curso, "&_
	  " (select case when bq.id_sede<>27 then ISNULL((select cm.NOMBRE from comuna cm "&_
	  " where cm.ID_COMUNA=s.ID_COMUNA),'') "&_
	  " when bq.id_sede=27 then ISNULL((select cm.NOMBRE from comuna cm  "&_
	  " where cm.ID_COMUNA=bq.id_comuna),'') end from bloque_programacion bq   "&_
	  " inner join sedes s on s.ID_SEDE=bq.id_sede "&_
	  " where bq.id_programa=PROGRAMA.ID_PROGRAMA) as 'Comuna',  "&_
	  " (select case when bq.id_sede<>27 then ISNULL((select r.NOMBRE from regiones r  "&_
	  " where r.ID_REGION=s.ID_REGION),'') "&_
	  " when bq.id_sede=27 then ISNULL((select r.NOMBRE from regiones r  "&_
	  " where r.ID_REGION=bq.id_region),'') end from bloque_programacion bq   "&_
	  " inner join sedes s on s.ID_SEDE=bq.id_sede "&_
	  " where bq.id_programa=PROGRAMA.ID_PROGRAMA) as 'Region',  "&_
	  " (select case when bq.id_sede<>27 then ISNULL((select r.id_region from regiones r  "&_
	  " where r.ID_REGION=s.ID_REGION),'') "&_
	  " when bq.id_sede=27 then ISNULL((select r.id_region from regiones r  "&_
	  " where r.ID_REGION=bq.id_region),'') end from bloque_programacion bq   "&_
	  " inner join sedes s on s.ID_SEDE=bq.id_sede "&_
	  " where bq.id_programa=PROGRAMA.ID_PROGRAMA) as 'idRegion',  "&_	  
	  " (select s.ID_COD_INT_MUTUAL from bloque_programacion bq  "&_
	  " inner join sedes s on s.ID_SEDE=bq.id_sede "&_
	  " where bq.id_programa=PROGRAMA.ID_PROGRAMA) as id_sede_curso, "&_
	  " HI=(select hb.HORARIO from HORARIO_BLOQUES hb where hb.ID_HORARIO=PROGRAMA.BMI),  "&_
	  " HF=(select hb.HORARIO from HORARIO_BLOQUES hb where hb.ID_HORARIO=PROGRAMA.BTF), "&_
	  " (select Ir.RUT from bloque_programacion bq  "&_
	  " inner join INSTRUCTOR_RELATOR ir on ir.ID_INSTRUCTOR=bq.id_relator  "&_
	  " where bq.id_programa=PROGRAMA.ID_PROGRAMA) as Rut_instructor,  "&_
	  " (select UPPER(Ir.NOMBRES+' '+Ir.A_PATERNO+' '+Ir.A_MATERNO) from bloque_programacion bq  "&_
	  " inner join INSTRUCTOR_RELATOR ir on ir.ID_INSTRUCTOR=bq.id_relator  "&_
	  " where bq.id_programa=PROGRAMA.ID_PROGRAMA) as instructor,CURRICULO.ID_PROYECTO, "&_
	  " (CASE CURRICULO.ID_PROYECTO WHEN 1 THEN 'Abierta' WHEN 2 THEN 'Cerrada' WHEN 3 THEN 'Proyecto Almagro' "&_
	  " WHEN 4 THEN 'Spot Mutual' WHEN 5 THEN 'Cerrado-Collahuasi' ELSE '' end) as 'PROYECTO', "&_
	  " (select case bq.estado_eva_cdn when 1 then 'Cerrado' when 2 then 'Cerrado Sin Asistentes' "&_
	  " when 3 then 'Cerrado Sin Participantes' when 4 then 'Curso Suspendido' when 5 then 'Curso Suspendido Sin Participantes' when 6 then 'Ejecutado por Mutual' "&_
      " else 'No Evaluado' end from bloque_programacion bq where bq.id_programa=PROGRAMA.ID_PROGRAMA) as 'EC',"&_ 
      "usr=ISNULL((select u.NOMBRES+' '+u.A_PATERNO+' '+u.A_MATERNO from usuarios u where u.ID_USUARIO=PROGRAMA.USUARIO_CREA),''), "&_	   
	  " PROGRAMA.ESTADO,CURRICULO.PGP from PROGRAMA  "&_
	  " inner join CURRICULO on CURRICULO.ID_MUTUAL=PROGRAMA.ID_MUTUAL "&_
	  " where PROGRAMA.ESTADO=1 AND PROGRAMA.ID_CLIENTE in (11) and programa.id_programa < 22232 "&proSel&mesSel&anoSel&_
	  " order by programa.FECHA_INICIO_"
	  '" where PROGRAMA.FECHA_INICIO_>=CONVERT(date,'"&request("f_ini")&"') "&_
	  '" and PROGRAMA.FECHA_INICIO_<DATEADD(DAY, 1, CONVERT(date,'"&request("f_fin")&"', 105))"
	  


	set rsMan =  conn.execute(qsMan)%>

    <table width="800" border="0">
        <tr>
          <td width="60" align="center" style="border-width: 1px;border: solid; border-color:#000;"><b><font size="2">ID</font></b></td>
		  <td width="150" align="center" style="border-width: 1px;border: solid; border-color:#000;"><b><font size="2">CODIGO CURSO</font></b></td>      
          <td width="850" align="center" style="border-width: 1px;border: solid; border-color:#000;"><b><font size="2">NOMBRE CURSO</font></b></td>
		  <td width="150" align="center" style="border-width: 1px;border: solid; border-color:#000;"><b><font size="2">FECHA INICIO</font></b></td>           		
          <td width="150" align="center" style="border-width: 1px;border: solid; border-color:#000;"><b><font size="2">FECHA TERMINO</font></b></td>
		  <td width="130" align="center" style="border-width: 1px;border: solid; border-color:#000;"><b><font size="2">CUPOS</font></b></td>
		 <td width="130" align="center" style="border-width: 1px;border: solid; border-color:#000;"><b><font size="2">INSCRITOS</font></b></td>                     
          <td width="130" align="center" style="border-width: 1px;border: solid; border-color:#000;"><b><font size="2">ASISTENTES</font></b></td>
		  <td width="130" align="center" style="border-width: 1px;border: solid; border-color:#000;"><b><font size="2">APROBADOS</font></b></td>           
          <td width="130" align="center" style="border-width: 1px;border: solid; border-color:#000;"><b><font size="2">REPROBADOS</font></b></td>
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
        </tr>
        <%
        while not rsMan.eof
        %>
        <tr>
          <td style="border-width: 1px;border: solid; border-color:#000;"><font size="2"><%=rsMan("ID_PROGRAMA")%></font></td>
          <td style="border-width: 1px;border: solid; border-color:#000;"><font size="2"><%=rsMan("CODIGO")%></font></td>          
          <td style="border-width: 1px;border: solid; border-color:#000;"><font size="2"><%=rsMan("NOMBRE_CURSO")%></font></td>
          <td align="center" style="border-width: 1px;border: solid; border-color:#000;"><font size="2"><%=rsMan("FECHA_INICIO")%></font></td>  
          <td align="center" style="border-width: 1px;border: solid; border-color:#000;"><font size="2"><%=rsMan("FECHA_TERMINO")%></font></td>
	      <td style="border-width: 1px;border: solid; border-color:#000;"><font size="2"><%=rsMan("CUPOS")%></font></td>	
          <td style="border-width: 1px;border: solid; border-color:#000;"><font size="2"><%=rsMan("inscritos")%></font></td>          
          <td style="border-width: 1px;border: solid; border-color:#000;"><font size="2"><%=rsMan("Asistentes")%></font></td>
          <td style="border-width: 1px;border: solid; border-color:#000;"><font size="2"><%=rsMan("Aprobados")%></font></td>  
          <td style="border-width: 1px;border: solid; border-color:#000;"><font size="2"><%=rsMan("Reprobados")%></font></td>
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
          <%if((rsMan("idRegion")=1 or rsMan("idRegion")=2 or rsMan("idRegion")=3 or rsMan("idRegion")=4 or rsMan("idRegion")=15) and rsMan("PROYECTO")="Abierta")then%>
          	<td style="border-width: 1px;border: solid; border-color:#000;"><font size="2">Ejecutado por Mutual</font></td>
          <%else%>
        	<td style="border-width: 1px;border: solid; border-color:#000;"><font size="2"><%=rsMan("EC")%></font></td>	
          <%end if%>
            	
          <td style="border-width: 1px;border: solid; border-color:#000;"><font size="2"><%=rsMan("usr")%></font></td>	
        </tr>	
        <%
            rsMan.MoveNext
        wend
        %>
    </table>
