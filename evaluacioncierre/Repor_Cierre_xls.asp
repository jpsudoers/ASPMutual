
<!--#include file="../conexion.asp"-->
<%
Response.CodePage = 65001
Response.CharSet = "utf-8"

fecha = year(now)&right("0"&month(now()),2)&right("0"&day(now()),2)
fecha = fecha&right("0"&hour(now()),2)&right("0"&minute(now()),2)&right("0"&second(now()),2)


Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=Reporte_Cierre_Curso_"&fecha&".xls"

sqlProg = "select (case when P.FECHA_INICIO_<P.FECHA_TERMINO then "&_
					"CONVERT(varchar(2), DAY(P.FECHA_INICIO_))+' '+CONVERT(varchar(3), (CASE MONTH(P.FECHA_INICIO_) "&_
					"WHEN 1 THEN 'Ene' "&_
					"WHEN 2 THEN 'Feb' "&_
					"WHEN 3 THEN 'Mar' "&_
					"WHEN 4 THEN 'Abr' "&_
					"WHEN 5 THEN 'May' "&_
					"WHEN 6 THEN 'Jun' "&_
					"WHEN 7 THEN 'Jul' "&_
					"WHEN 8 THEN 'Ago' "&_
					"WHEN 9 THEN 'Sep' "&_
					"WHEN 10 THEN 'Oct' "&_
					"WHEN 11 THEN 'Nov' "&_
					"WHEN 12 THEN 'Dic' END))+' '+ CONVERT(varchar(4), Year(P.FECHA_INICIO_))+' y '+  "&_
					"CONVERT(varchar(2), DAY(P.FECHA_TERMINO))+' '+CONVERT(varchar(3), (CASE MONTH(P.FECHA_TERMINO) "&_
					"WHEN 1 THEN 'Ene' "&_
					"WHEN 2 THEN 'Feb' "&_
					"WHEN 3 THEN 'Mar' "&_
					"WHEN 4 THEN 'Abr' "&_
					"WHEN 5 THEN 'May' "&_
					"WHEN 6 THEN 'Jun' "&_
					"WHEN 7 THEN 'Jul' "&_
					"WHEN 8 THEN 'Ago' "&_
					"WHEN 9 THEN 'Sep' "&_
					"WHEN 10 THEN 'Oct' "&_
					"WHEN 11 THEN 'Nov' "&_
					"WHEN 12 THEN 'Dic' "&_
					"END))+' '+ CONVERT(varchar(4), Year(P.FECHA_TERMINO) ) else "&_
					"CONVERT(varchar(2), DAY(P.FECHA_INICIO_))+' '+CONVERT(varchar(3), (CASE MONTH(P.FECHA_INICIO_) "&_
					"WHEN 1 THEN 'Ene' "&_
					"WHEN 2 THEN 'Feb' "&_
					"WHEN 3 THEN 'Mar' "&_
					"WHEN 4 THEN 'Abr' "&_
					"WHEN 5 THEN 'May' "&_
					"WHEN 6 THEN 'Jun' "&_
					"WHEN 7 THEN 'Jul' "&_
					"WHEN 8 THEN 'Ago' "&_
					"WHEN 9 THEN 'Sep' "&_
					"WHEN 10 THEN 'Oct' "&_
					"WHEN 11 THEN 'Nov' "&_
					"WHEN 12 THEN 'Dic' END))+' '+ CONVERT(varchar(4), Year(P.FECHA_INICIO_)) "&_
					"end) as fecha,p.DIR_EJEC,s.CIUDAD from bloque_programacion bq  "&_
					"inner join PROGRAMA p on P.ID_PROGRAMA=bq.id_programa   "&_
					"inner join sedes s on s.ID_SEDE=bq.id_sede "&_
					"where bq.ID_BLOQUE="&Request("b")

					set rsProg = conn.execute (sqlProg)
					%>
<table>
    <tr>
        <td>REPORTE DE INDUCCI�N</td>
    </tr>
</table>
<table border="1">
    <tr>
        <td>N�</td>
        <td>EMPRESA</td>
        <td>PROYECTO</td>
        <td>NOMBRE</td>
        <td>RUT</td>
        <td>FECHA REALIZACI�N</td>
        <td>ESTATUS D�A I</td>
        <td>ESTATUS D�A II</td>
        <td>CIUDAD</td>
  	</tr>
	
	<%
	sql = "select dbo.MayMinTexto(T.NOMBRES) as NOMBRES,"&_
	         "(CASE WHEN T.NACIONALIDAD='1' then T.ID_EXTRANJERO WHEN T.NACIONALIDAD='0' then T.RUT END) as 'TrabId',"&_
			 " hc.ASISTENCIA,hc.CALIFICACION,hc.EVALUACION, "&_
			 " hc.ASIS_CDN,hc.CAL_CDN,hc.EVA_CDN, "&_
			 " hc.ASIS_REL,hc.CAL_REL,hc.EVA_REL,dbo.MayMinTexto(e.R_SOCIAL) as R_SOCIAL,"&_
    		 " CONVERT(VARCHAR(10),P.FECHA_INICIO_, 105) as FECHA_INICIO_,"&_
    		 " CONVERT(VARCHAR(10),P.FECHA_TERMINO, 105) as FECHA_TERMINO "&_
			 " from HISTORICO_CURSOS hc "&_
			 " inner join trabajador t on t.ID_TRABAJADOR=hc.ID_TRABAJADOR "&_
			 " inner join bloque_programacion bq on bq.id_bloque=hc.ID_BLOQUE "&_
             " inner join PROGRAMA p on P.ID_PROGRAMA=bq.id_programa "&_ 
     	     " inner join empresas e on e.ID_EMPRESA=hc.ID_EMPRESA "&_		 
			 " where bq.ID_BLOQUE="&Request("b")&_
			 " ORDER BY R_SOCIAL, NOMBRES ASC"

	set rsTrab = conn.execute (sql)

	cont=0
	n_trab=1
	
	while not rsTrab.eof
				cont=cont+1
				%>
	<tr>
        <td><%=n_trab %></td>
        <td><%=rsTrab("R_SOCIAL") %></td>
        <td></td>
        <td><%=rsTrab("NOMBRES") %></td>
        <td><%=rsTrab("TrabId") %></td>
        <td><%=rsProg("fecha") %></td>
        <td></td>
        <td></td>
        <td><%=rsProg("CIUDAD") %></td>
  	</tr>											
					<%			
				n_trab=n_trab+1
			rsTrab.Movenext
	wend	
	
	




%>  

</table>