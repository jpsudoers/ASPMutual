<!--#include file="../conexion.asp"-->
<%
fecha = year(now)&right("0"&month(now()),2)&right("0"&day(now()),2)
fecha = fecha&right("0"&hour(now()),2)&right("0"&minute(now()),2)&right("0"&second(now()),2)
Response.ContentType ="application/vnd.ms-excel"
Response.AddHeader "content-disposition", "inline; filename=XLSCertificados_"&fecha&".xls"
	
	'insSel=""
	'if(request("insc")="0" or request("insc")="1")then
	'	  insSel=" AUTORIZACION.estado="&request("insc")
	'elseif(request("insc")="2")then
	'	  insSel=" AUTORIZACION.estado=0 and AUTORIZACION.facturado=1 "
	'else
	'	  insSel=" AUTORIZACION.estado in (0,1)"
	'end if
	
	
	
	qsHc="select dbo.MayMinTexto(E.R_SOCIAL) as R_SOCIAL,T.RUT,dbo.MayMinTexto(T.NOMBRES) as NOMBRES,"&_ 
		" dbo.MayMinTexto(C.NOMBRE_CURSO) as NOMBRE_CURSO,CONVERT(VARCHAR(10),P.FECHA_INICIO_, 105) as FECHA_INICIO, "&_ 
		" E.ID_EMPRESA,H.ID_PROGRAMA,H.RELATOR,H.ID_TRABAJADOR,P.FECHA_INICIO_ "&_ 
		" from HISTORICO_CURSOS H "&_ 
		" inner join EMPRESAS E on E.ID_EMPRESA=H.ID_EMPRESA "&_ 
		" inner join TRABAJADOR T on T.ID_TRABAJADOR=H.ID_TRABAJADOR "&_ 
		" inner join PROGRAMA P on P.ID_PROGRAMA=H.ID_PROGRAMA "&_ 
		" inner join CURRICULO C on C.ID_MUTUAL=P.ID_MUTUAL "&_ 
		" where H.EVALUACION<>'Reprobado' "&_ 
		" and H.ESTADO=0 and e.ID_EMPRESA in ('"&request("e")&"')" 
		
		
		if(request("t") <> "") then
			qsHc = qsHc & " and T.ID_TRABAJADOR = ('"&request("t")&"')"
		end if	
		
		qsHc = qsHc & " order by P.FECHA_INICIO_,NOMBRES"

	set rsHc =  conn.execute(qsHc)%>

    <table width="2250" border="1">
    <tr>
      <td width="550" align="center"><b><font size="2">Razón Social</font></b></td>  
      <td width="150" align="center"><b><font size="2">Fecha Curso</font></b></td>        
      <td width="150" align="center"><b><font size="2">Rut Trabajador</font></b></td>       
      <td width="550" align="center"><b><font size="2">Nombre Trabajador</font></b></td>       
      <td width="550" align="center"><b><font size="2">Curso</font></b></td>
      <td width="100" align="center"><b><font size="2">Ver Certificado</font></b></td>   
    </tr>
	<%
	nomArch=""
	aliasArch=""
    while not rsHc.eof
		'if request("c")="82" then
			'aliasArch="CertificadoCodelco"
		'else
                if request("c")="52" then
			aliasArch="CertificadoCMCC"	
		else
			aliasArch="Certificado"	
		end if
	
	nomArch="http://norte.otecmutual.cl/libroclases/"&aliasArch&".asp?prog="&rsHc("ID_PROGRAMA")&"&empresa="&rsHc("ID_EMPRESA")&"&trabajador="&rsHc("ID_TRABAJADOR")&"&relator="&rsHc("RELATOR")
    %>
    <tr>
      <td><font size="2"><%=rsHc("R_SOCIAL")%></font></td>
      <td><font size="2"><%=rsHc("FECHA_INICIO")%></font></td>      
      <td><font size="2"><%=rsHc("RUT")%></font></td>      
      <td><font size="2"><%=rsHc("NOMBRES")%></font></td>     
      <td><font size="2"><%=rsHc("NOMBRE_CURSO")%></font></td>
      <td align="center"><b><font size="2"><a href="<%=nomArch%>">Ver</a></font></b></td>  
    </tr>	
    <%
        rsHc.MoveNext
    wend
    %>
    </table>
