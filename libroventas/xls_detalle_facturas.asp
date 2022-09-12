<!--#include file="../conexion.asp"-->
<%
fecha = year(now)&right("0"&month(now()),2)&right("0"&day(now()),2)
fecha = fecha&right("0"&hour(now()),2)&right("0"&minute(now()),2)&right("0"&second(now()),2)



txe=""
if(Request("e")<>"" and Request("e")<>"0")then
	txe=" and empresas.id_empresa='"&Request("e")&"'"
end if

txm=""
if(Request("m")<>"" and Request("m")<>"0")then
	txm=" and month(FACTURAS.FECHA_EMISION)="&Request("m")
end if

txa=""
if(Request("a")<>"")then
	txa=" and year(FACTURAS.FECHA_EMISION)="&Request("a")
end if

qs="select 'RUT'=empresas.RUT,"&_
	"'NOMBRE AUXILIAR'=UPPER (empresas.R_SOCIAL),"&_
	"'N° DOCTO.'=facturas.FACTURA,"&_
	"'F. EMISION'=CONVERT(VARCHAR(10),facturas.FECHA_EMISION, 105),"&_
	"'F. VCTO.'=CONVERT(VARCHAR(10),facturas.FECHA_VENCIMIENTO, 105), "&_
	"DEBE='$ ' + CONVERT(VARCHAR(1000), facturas.MONTO),"&_
	"HABER='$ - ',"&_
	"SALDO='$ ' + CONVERT(VARCHAR(1000), facturas.MONTO),"&_
	"'ORDEN DE COMPRA'=(CASE AUTORIZACION.CON_OTIC WHEN 0 THEN AUTORIZACION.ORDEN_COMPRA ELSE '' END),"&_
	"'O-C OTIC'=(CASE AUTORIZACION.CON_OTIC WHEN 1 THEN AUTORIZACION.ORDEN_COMPRA ELSE '' END),"&_
	"'PARTICIPANTE'=UPPER (trabajador.NOMBRES), "&_
	"'RUT PARTICIPANTE'=(CASE WHEN trabajador.NACIONALIDAD=0 then trabajador.RUT  "&_
	" WHEN trabajador.NACIONALIDAD=1 then trabajador.ID_EXTRANJERO END),"&_
	" Estado=HISTORICO_CURSOS.EVALUACION,"&_
	"Fecha=CONVERT(VARCHAR(10),PROGRAMA.FECHA_INICIO_, 105),"&_
	" XXXX=(CASE AUTORIZACION.CON_OTIC WHEN 1 THEN '0' ELSE '1' END),facturas.FECHA_EMISION"&_
	" from HISTORICO_CURSOS  "&_
	" inner join trabajador on trabajador.ID_TRABAJADOR=HISTORICO_CURSOS.ID_TRABAJADOR  "&_
	" inner join empresas on EMPRESAS.ID_EMPRESA=HISTORICO_CURSOS.ID_EMPRESA  "&_
	" inner join PROGRAMA on programa.ID_PROGRAMA=HISTORICO_CURSOS.ID_PROGRAMA  "&_
	" inner join AUTORIZACION on AUTORIZACION.ID_AUTORIZACION=HISTORICO_CURSOS.ID_AUTORIZACION 	"&_
	" inner join facturas on facturas.ID_AUTORIZACION=AUTORIZACION.ID_AUTORIZACION "&_
	" where HISTORICO_CURSOS.CALIFICACION is not null and facturas.FACTURA is not null "&txe&txm&txa&_
	" --and CURRICULO.ID_CLIENTE in (1) "&_
	" --AND PROGRAMA.FECHA_INICIO_>=convert(date, '01-11-2013') "&_	
	" --and empresas.RUT='91915000-9' "&_
	" order by facturas.FECHA_EMISION,EMPRESAS.R_SOCIAL,TRABAJADOR.NOM_TRAB"

	N_Doc="0"
	set rs =  conn.execute(qs)
	
'	Response.Clear()
'Response.ContentType ="application/vnd.ms-excel"
response.contentType = "application/vnd.ms-excel; charset=utf-8"
Response.AddHeader "content-disposition", "inline; filename=Detalle_Facturas_"&fecha&".xls"
Session.CodePage = 65001
Response.charset ="utf-8"
'response.charset = "utf-8"
	%>
    <table border="1">
    <tr>
      <td width="120" align="center"><b><font size="2">RUT</font></b></td>
      <td width="600" align="center"><b><font size="2">NOMBRE AUXILIAR</font></b></td>
      <td width="130" align="center"><b><font size="2">N&quot; DOCTO.</font></b></td>
      <td width="130" align="center"><b><font size="2">F. EMISION</font></b></td>  
      <td width="130" align="center"><b><font size="2">F. VCTO.</font></b></td>
      <td width="130" align="center"><b><font size="2">DEBE</font></b></td>    
      <td width="130" align="center"><b><font size="2">HABER</font></b></td> 
      <td width="130" align="center"><b><font size="2">SALDO</font></b></td>           
      <td width="130" align="center"><b><font size="2">ORDEN DE COMPRA</font></b></td>         
      <td width="130" align="center"><b><font size="2">O-C OTIC</font></b></td>  
      <td width="350" align="center"><b><font size="2">PARTICIPANTE</font></b></td>     
      <td width="130" align="center"><b><font size="2">RUT PARTICIPANTE</font></b></td>    
      <td width="130" align="center"><b><font size="2">Estado</font></b></td>  
      <td width="130" align="center"><b><font size="2">Fecha</font></b></td>     
      <td width="130" align="center"><b><font size="2">XXXX</font></b></td>   
    </tr>
	<%
    while not rs.eof
	%>
    	<tr>
    <%
	 if(N_Doc<>rs("N° DOCTO."))then
    %>
          <td><font size="2"><%=rs("RUT")%></font></td>
          <td><font size="2"><%=rs("NOMBRE AUXILIAR")%></font></td>
          <td><font size="2"><%=rs("N° DOCTO.")%></font></td>          
          <td><font size="2"><%=rs("F. EMISION")%></font></td>
          <td><font size="2"><%=rs("F. VCTO.")%></font></td>    
          <td><font size="2"><%=rs("DEBE")%></font></td>  
          <td><font size="2"><%=rs("HABER")%></font></td>  
          <td><font size="2"><%=rs("SALDO")%></font></td>                
          <td><font size="2"><%=rs("ORDEN DE COMPRA")%></font></td>  
          <td><font size="2"><%=rs("O-C OTIC")%></font></td>
     <%
	 else
	 %>     
          <td></td>
          <td></td>
          <td></td>          
          <td></td>
          <td></td>    
          <td></td>  
          <td></td>  
          <td></td>                
          <td></td>  
          <td></td>
     <%
	 end if
	 %>
          <td><font size="2"><%=rs("PARTICIPANTE")%></font></td>  
          <td><font size="2"><%=rs("RUT PARTICIPANTE")%></font></td>    
          <td><font size="2"><%=rs("Estado")%></font></td>  
          <td><font size="2"><%=rs("Fecha")%></font></td>          
          <td><font size="2"><%=rs("XXXX")%></font></td>                                          
        </tr>
		<%
		N_Doc=rs("N° DOCTO.")
	rs.MoveNext
wend
%>
</table>