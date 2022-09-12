<!--#include file="../conexion.asp"-->
<%
Response.ContentType ="application/vnd.ms-excel; utf-8"
'Response.AddHeader "content-disposition", "inline; filename=REPOR_FACT_"&fecha&".xls"
txe=""
if(Request("e")<>"" and Request("e")<>"0")then
	txe=" and E.id_empresa='"&Request("e")&"'"
end if

txm=""
if(Request("m")<>"" and Request("m")<>"0")then
	txm=" and month(F.FECHA_EMISION)="&Request("m")
end if

txa=""
if(Request("a")<>"")then
	txa=" and year(F.FECHA_EMISION)="&Request("a")
end if
sqlempresa = "select top 1 RUT from EMPRESAS where id_empresa="&Request("e")
set resRut = conn.execute(sqlempresa)
'qs="select 'RUT'=empresas.RUT,"&_
'	"'NOMBRE AUXILIAR'=UPPER (empresas.R_SOCIAL),"&_
'	"'N° DOCTO.'=facturas.FACTURA,"&_
'	"'F. EMISION'=CONVERT(VARCHAR(10),facturas.FECHA_EMISION, 105),"&_
'	"'F. VCTO.'=CONVERT(VARCHAR(10),facturas.FECHA_VENCIMIENTO, 105), "&_
'	"DEBE='$ ' + CONVERT(VARCHAR(1000), facturas.MONTO),"&_
'	"HABER='$ - ',"&_
'	"SALDO='$ ' + CONVERT(VARCHAR(1000), facturas.MONTO),"&_
'	"'ORDEN DE COMPRA'=(CASE AUTORIZACION.CON_OTIC WHEN 0 THEN AUTORIZACION.ORDEN_COMPRA ELSE '' END),"&_
'	"'O-C OTIC'=(CASE AUTORIZACION.CON_OTIC WHEN 1 THEN AUTORIZACION.ORDEN_COMPRA ELSE '' END),"&_
'	"'PARTICIPANTE'=UPPER (trabajador.NOMBRES), "&_
'	"'RUT PARTICIPANTE'=(CASE WHEN trabajador.NACIONALIDAD=0 then trabajador.RUT  "&_
'	" WHEN trabajador.NACIONALIDAD=1 then trabajador.ID_EXTRANJERO END),"&_
'	" Estado=HISTORICO_CURSOS.EVALUACION,"&_
'	"Fecha=CONVERT(VARCHAR(10),PROGRAMA.FECHA_INICIO_, 105),"&_
'	" XXXX=(CASE AUTORIZACION.CON_OTIC WHEN 1 THEN '0' ELSE '1' END),facturas.FECHA_EMISION"&_
'	" from HISTORICO_CURSOS  "&_
'	" inner join trabajador on trabajador.ID_TRABAJADOR=HISTORICO_CURSOS.ID_TRABAJADOR  "&_
'	" inner join empresas on EMPRESAS.ID_EMPRESA=HISTORICO_CURSOS.ID_EMPRESA  "&_
'	" inner join PROGRAMA on programa.ID_PROGRAMA=HISTORICO_CURSOS.ID_PROGRAMA  "&_
'	" inner join AUTORIZACION on AUTORIZACION.ID_AUTORIZACION=HISTORICO_CURSOS.ID_AUTORIZACION 	"&_
'	" inner join facturas on facturas.ID_AUTORIZACION=AUTORIZACION.ID_AUTORIZACION "&_
'	" where HISTORICO_CURSOS.CALIFICACION is not null and facturas.FACTURA is not null "&txe&txm&txa&_
'	" --and CURRICULO.ID_CLIENTE in (1) "&_
'	" --AND PROGRAMA.FECHA_INICIO_>=convert(date, '01-11-2013') "&_	
'	" --and empresas.RUT='91915000-9' "&_
'	" order by facturas.FECHA_EMISION,EMPRESAS.R_SOCIAL,TRABAJADOR.NOM_TRAB"
qs="SELECT c.DESCRIPCION as Curso, a.N_PARTICIPANTES as Participantes, a.VALOR_OC as Valor FROM FACTURAS F "&_ 
   "inner join AUTORIZACION A on A.ID_AUTORIZACION=F.ID_AUTORIZACION "&_ 
   "inner join EMPRESAS EI ON EI.ID_EMPRESA=A.ID_EMPRESA "&_   
   "inner join EMPRESAS E on E.ID_EMPRESA=F.ID_EMPRESA "&_ 
   "inner join PROGRAMA P on P.ID_PROGRAMA=A.ID_PROGRAMA "&_   
   "inner join CURRICULO C on C.ID_MUTUAL=P.ID_MUTUAL "&_ 
   "inner join bloque_programacion BQ on BQ.id_bloque=A.ID_BLOQUE "&_ 
   "inner join  EMPRESA_CONDICION_COMERCIAL cc on E.ID_EMPRESA = cc.ID_EMPRESA WHERE A.ESTADO in (0,1) "&_  
   "and  cc.ID_CONDICION_COMERCIAL in (0,5) and F.FECHA_PAGO is null and F.DESCRIPCION_ESTADO is null "&txe&txa&txm
   
   
	N_Doc="0"
	set rs =  conn.execute(qs)
	total = 0
'	Response.Clear()
'Response.ContentType ="application/vnd.ms-excel"
response.contentType = "application/vnd.ms-excel; charset=utf-8"
Response.AddHeader "content-disposition", "inline; filename=Detalle_EDP_"&resRut("RUT")&".xls"
Session.CodePage = 65001
Response.charset ="utf-8"
'response.charset = "utf-8"
	%>
    <table border="1">
    <tr>
      <td width="120" align="center"><b><font size="2">Curso</font></b></td>
      <td width="120" align="center"><b><font size="2">PARTICIPANTE</font></b></td>
      <td width="130" align="center"><b><font size="2">Valor</font></b></td>

    </tr>
	<%
    while not rs.eof
	%>
    	<tr>

          <td><font size="2"><%=rs("Curso")%></font></td>
          <td><font size="2"><%=rs("Participantes")%></font></td>
          <td><font size="2"><%=rs("Valor")%></font></td>          
                                      
        </tr>
		<%
	total = total + rs("Valor")
	rs.MoveNext
wend
%>
		<tr>

          <td><font size="2"></font></td>
          <td><font size="2"></font></td>
          <td><font size="2"><%=total%></font></td>          
                                      
        </tr>
</table>