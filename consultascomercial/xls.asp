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
			mesSel=" and Month(FECHA_INICIO_)="&request("mes")
		end if

		anoSel=""
		if(request("ano")<>"0")then
			anoSel=" and Year(FECHA_INICIO_)="&request("ano")
		end if	
	
		proSel=""
		if(request("p")<>"0")then
			proSel=" and ID_PROYECTO="&request("p")
		end if		

	conn.execute("update bloque_programacion set id_programa=0 where LEN(id_programa)>10")

	if(request("v")<>"2")then
        	qsMan= "SELECT * FROM dbmas.[dbo].VW_INFORME_COMERCIAL "&_
               	       " WHERE 1=1 "&mesSel&anoSel&" order by FECHA_INICIO_"
	else
		qsMan="SELECT * FROM dbarauco.[dbo].VW_INFORME_COMERCIAL "&_
            	      " WHERE 1=1 "&proSel&mesSel&anoSel&_
            	      " UNION ALL "&_
            	      " SELECT * FROM dbmas.[dbo].VW_INFORME_COMERCIAL "&_ 
                      " WHERE 1=1 "&mesSel&anoSel&" order by FECHA_INICIO_"
	end if  

	set rsMan =  conn.execute(qsMan)	

	%> 
    
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />    
    <table width="800" border="0">
    <tr>
    <td width="60" align="center" style="border-width: 1px;border: solid; border-color:#000;"><b><font size="2">ID</font></b></td>
    <td width="80" align="center" style="border-width: 1px;border: solid; border-color:#000;"><b><font size="2">Folio</font></b></td>
    <td width="80" align="center" style="border-width: 1px;border: solid; border-color:#000;"><b><font size="2">SIGECAP</font></b></td>
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
 		  <td width="350" align="center" style="border-width: 1px;border: solid; border-color:#000;"><b><font size="2">PROYECTOR</font></b></td>-->
 		  <td width="350" align="center" style="border-width: 1px;border: solid; border-color:#000;"><b><font size="2">ALMUERZO</font></b></td>
		  <td width="350" align="center" style="border-width: 1px;border: solid; border-color:#000;"><b><font size="2">MONITOR</font></b></td>
 		  <td width="350" align="center" style="border-width: 1px;border: solid; border-color:#000;"><b><font size="2">OTROS</font></b></td>  
 		  <td width="350" align="center" style="border-width: 1px;border: solid; border-color:#000;"><b><font size="2">DETALLE OTROS</font></b></td>         
          <td width="350" align="center" style="border-width: 1px;border: solid; border-color:#000;"><b><font size="2">ACT. ECONOMICA</font></b></td>
          <td width="350" align="center" style="border-width: 1px;border: solid; border-color:#000;"><b><font size="2">ACT. ECO. DESCRIPCION</font></b></td>
          <td width="350" align="center" style="border-width: 1px;border: solid; border-color:#000;"><b><font size="2">ACT. ECO. GRUPO</font></b></td>                                                                    
          <td width="350" align="center" style="border-width: 1px;border: solid; border-color:#000;"><b><font size="2">TIPO CURSO</font>    
          <td width="350" align="center" style="border-width: 1px;border: solid; border-color:#000;"><b><font size="2">ASISTENTES REPORTADOS</font>      
          <!--FIN costos del curso-->          
                                      
        </tr>
        <%
		
        while not rsMan.eof
        %>
        <tr>
          <td style="border-width: 1px;border: solid; border-color:#000;"><font size="2"><%=rsMan("ID_PROGRAMA")%></font></td>
          <td style="border-width: 1px;border: solid; border-color:#000;"><font size="2"><%=rsMan("NFOLIO")%></font></td>
          <td style="border-width: 1px;border: solid; border-color:#000;"><font size="2"><%=rsMan("ID_PLANIFICACION_SIGECAP")%></font></td>
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
          <td style="border-width: 1px;border: solid; border-color:#000;"><font size="2"><%=rsMan("EC")%></font></td>     	
          <td style="border-width: 1px;border: solid; border-color:#000;"><font size="2"><%=rsMan("usr")%></font></td>
          
  <!--INICIO costos del curso-->         
 		  <td style="border-width: 1px;border: solid; border-color:#000;"><font size="2"><%=rsMan("ARRIENDO_sala")%></font></b></td>          
 		  <td style="border-width: 1px;border: solid; border-color:#000;"><font size="2"><%=rsMan("RELATORIA")%></font></b></td>  
 		  <td style="border-width: 1px;border: solid; border-color:#000;"><font size="2"><%=rsMan("COFFEE")%></font></b></td>  
 		  <td style="border-width: 1px;border: solid; border-color:#000;"><font size="2"><%=rsMan("TRASLADO")%></font></b></td>  
 		  <!--<td style="border-width: 1px;border: solid; border-color:#000;"><font size="2"><%=rsMan("MATERIAL")%></font></b></td>  
 		  <td style="border-width: 1px;border: solid; border-color:#000;"><font size="2"><%=rsMan("FANTOMA")%></font></b></td>  
 		  <td style="border-width: 1px;border: solid; border-color:#000;"><font size="2"><%=rsMan("SIMULADOR")%></font></b></td>
 		  <td style="border-width: 1px;border: solid; border-color:#000;"><font size="2"><%=rsMan("PROYECTOR")%></font></b></td>-->
 		  <td style="border-width: 1px;border: solid; border-color:#000;"><font size="2"><%=rsMan("ALMUERZO")%></font></b></td>
		  <td style="border-width: 1px;border: solid; border-color:#000;"><font size="2"><%=rsMan("MONITOR")%></font></b></td>
 		  <td style="border-width: 1px;border: solid; border-color:#000;"><font size="2"><%=rsMan("OTROS")%></font></b></td>
  		  <td style="border-width: 1px;border: solid; border-color:#000;"><font size="2"><%=rsMan("Detalle_Otros")%></font></b></td>         
<td style="border-width: 1px;border: solid; border-color:#000;"><font size="2"><%=rsMan("ACT_ECONOMICA")%></font></b></td>
<td style="border-width: 1px;border: solid; border-color:#000;"><font size="2"><%=rsMan("ACT_ECO_DESCRIPCION")%></font></b></td>
<td style="border-width: 1px;border: solid; border-color:#000;"><font size="2"><%=rsMan("GRUPO_TEXTO")%></font></b></td>
<td style="border-width: 1px;border: solid; border-color:#000;"><font size="2"><%=rsMan("TIPO_INDUCCION_DESC")%></font></b></td>
		  <td style="border-width: 1px;border: solid; border-color:#000;"><font size="2"><%=rsMan("Asistentes Reportados")%></font></b></td>  
<!--FIN costos del curso-->        
          	
        </tr>	
        <%
            rsMan.MoveNext
        wend
        %>
    </table>
