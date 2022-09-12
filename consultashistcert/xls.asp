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
	Response.AddHeader "content-disposition", "inline; filename=Informes_Certificados_"&fecha&".xls"
	
qsMan="select lg.FECHA_GENERACION,rute=e.RUT,dbo.MayMinTexto(E.R_SOCIAL) as R_SOCIAL,"&_
	  "T.RUT,dbo.MayMinTexto(T.NOMBRES) as NOMBRES,C.NOMBRE_CURSO,"&_
	  "CONVERT(VARCHAR(10),P.FECHA_TERMINO, 105) as FECHA_TERMINO,"&_
	  "CONVERT(VARCHAR(10),DATEADD(MONTH, (CASE "&_
	  " WHEN C.VIGENCIA =  1 THEN 6 "&_
	  " WHEN C.VIGENCIA =  2 THEN 12 "&_
	  " WHEN C.VIGENCIA =  3 THEN 18 "&_
	  " WHEN C.VIGENCIA =  4 THEN 24 "&_
	  " WHEN C.VIGENCIA =  5 THEN 48 "&_
	  " END), CONVERT(date,DATEADD(DAY, 1, CONVERT(date,P.FECHA_TERMINO, 105)), 105)), 105) as vigencia,"&_
	  " (CASE WHEN lg.GENERADO=1 then 'Si' WHEN lg.GENERADO=2 then 'Consultado' else 'No' END) as 'Gen',lg.TIPO_ERROR "&_
	  "  from LOG_CERTIFICADOS lg "&_
	  " inner join EMPRESAS E on E.ID_EMPRESA=lg.ID_EMPRESA "&_
	  " inner join TRABAJADOR T on T.ID_TRABAJADOR=lg.ID_TRABAJADOR "&_
	  " inner join PROGRAMA P on P.ID_PROGRAMA=lg.ID_PROGRAMA "&_
	  " inner join CURRICULO C on C.ID_MUTUAL=P.ID_MUTUAL "&_
	  " where lg.FECHA_GENERACION>=CONVERT(date,'"&request("f_ini")&"') "&_
	  " and lg.FECHA_GENERACION<DATEADD(DAY, 1, CONVERT(date,'"&request("f_fin")&"', 105))"
	

	set rsMan =  conn.execute(qsMan)%>

    <table width="800" border="0">
        <tr>
          <td width="150" align="center" style="border-width: 1px;border: solid; border-color:#000;"><b><font size="2">Fecha y Hora</font></b></td>
		  <td width="100" align="center" style="border-width: 1px;border: solid; border-color:#000;"><b><font size="2">Rut Empresa</font></b></td>      
          <td width="500" align="center" style="border-width: 1px;border: solid; border-color:#000;"><b><font size="2">Nombre Empresa</font></b></td>
		  <td width="100" align="center" style="border-width: 1px;border: solid; border-color:#000;"><b><font size="2">Rut Participante</font></b></td>           		
          <td width="500" align="center" style="border-width: 1px;border: solid; border-color:#000;"><b><font size="2">Nombre Participante</font></b></td>
		  <td width="200" align="center" style="border-width: 1px;border: solid; border-color:#000;"><b><font size="2">Proyecto</font></b></td>
		 <td width="300" align="center" style="border-width: 1px;border: solid; border-color:#000;"><b><font size="2">Curso</font></b></td>                     
          <td width="100" align="center" style="border-width: 1px;border: solid; border-color:#000;"><b><font size="2">Fecha Curso</font></b></td>
		  <td width="100" align="center" style="border-width: 1px;border: solid; border-color:#000;"><b><font size="2">Vigencia Curso</font></b></td>           
          <td width="100" align="center" style="border-width: 1px;border: solid; border-color:#000;"><b><font size="2">Certificado Generado</font></b></td>
		  <td width="300" align="center" style="border-width: 1px;border: solid; border-color:#000;"><b><font size="2">Tipo de Error</font></b></td>               
        </tr>
        <%
        while not rsMan.eof
        %>
        <tr>
          <td style="border-width: 1px;border: solid; border-color:#000;"><b><font size="2"><%=rsMan("FECHA_GENERACION")%></font></b></td>
          <td style="border-width: 1px;border: solid; border-color:#000;"><b><font size="2"><%=rsMan("rute")%></font></b></td>          
          <td style="border-width: 1px;border: solid; border-color:#000;"><b><font size="2"><%=rsMan("R_SOCIAL")%></font></b></td>
          <td style="border-width: 1px;border: solid; border-color:#000;"><b><font size="2"><%=rsMan("RUT")%></font></b></td>  
          <td style="border-width: 1px;border: solid; border-color:#000;"><b><font size="2"><%=rsMan("NOMBRES")%></font></b></td>
		  <td style="border-width: 1px;border: solid; border-color:#000;"><b><font size="2"></font></b></td>           
		  <td style="border-width: 1px;border: solid; border-color:#000;"><b><font size="2"><%=rsMan("NOMBRE_CURSO")%></font></b></td>		  
          <td style="border-width: 1px;border: solid; border-color:#000;"><b><font size="2"><%=rsMan("FECHA_TERMINO")%></font></b></td>  
          <td style="border-width: 1px;border: solid; border-color:#000;"><b><font size="2"><%=rsMan("vigencia")%></font></b></td>
          <td style="border-width: 1px;border: solid; border-color:#000;"><b><font size="2"><%=rsMan("Gen")%></font></b></td>  
          <td style="border-width: 1px;border: solid; border-color:#000;"><b><font size="2"><%=rsMan("TIPO_ERROR")%></font></b></td>
        </tr>	
        <%
            rsMan.MoveNext
        wend
        %>
    </table>
