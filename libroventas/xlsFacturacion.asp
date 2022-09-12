<!--#include file="../conexion.asp"-->
<%
fecha = year(now)&right("0"&month(now()),2)&right("0"&day(now()),2)
fecha = fecha&right("0"&hour(now()),2)&right("0"&minute(now()),2)&right("0"&second(now()),2)
Response.ContentType ="application/vnd.ms-excel"
Response.AddHeader "content-disposition", "inline; filename=REPOR_FACT_"&fecha&".xls"

bdoc=""
if(Request("txDoc")<>"")then
	bdoc=" and F.factura='"&Request("txDoc")&"'"
end if

txfcd=""
if(Request("fcd")<>"")then
	txfcd=" and P.FECHA_INICIO_>=CONVERT(DATE,'"&Request("fcd")&"')"
end if

txfch=""
if(Request("fch")<>"")then
	txfch=" and P.FECHA_INICIO_<=CONVERT(DATE,'"&Request("fch")&"')"
end if

txfed=""
if(Request("fed")<>"")then
	txfed=" and F.FECHA_EMISION>=CONVERT(DATE,'"&Request("fed")&"')"
end if

txfeh=""
if(Request("feh")<>"")then
	txfeh=" and F.FECHA_EMISION<=CONVERT(DATE,'"&Request("feh")&"')"
end if

txe=""
if(Request("e")<>"")then
	txe=" and F.ID_EMPRESA='"&Request("e")&"'"
end if

txc=""
if(Request("c")<>"0")then
	txc=" and C.ID_MUTUAL='"&Request("c")&"'"
end if

qs="SELECT E.RUT as rut,dbo.MayMinTexto (E.R_SOCIAL) as empresa,"&_
   "CONVERT(VARCHAR(10),P.FECHA_INICIO_, 105) as FECHA_INICIO,"&_
   "CONVERT(VARCHAR(10),P.FECHA_TERMINO, 105) as FECHA_TERMINO,"&_
   "C.CODIGO, dbo.MayMinTexto (C.NOMBRE_CURSO) as nombre, A.N_PARTICIPANTES, "&_
   "A.ID_AUTORIZACION,'http://norte.otecmutual.cl/ordenes/'+A.DOCUMENTO_COMPROMISO as doc,"&_
   "CONVERT(VARCHAR(10),F.FECHA_EMISION,105) as FECHA_EMISION,F.FACTURA,"&_
   "(CASE F.ESTADO WHEN 1 THEN 'Vigente' WHEN 0 THEN 'Anulada' WHEN 2 THEN 'Vig/Ref' end) as 'DESCEST',"&_
   "EI.RUT AS RUT_EMPRESA,EI.R_SOCIAL AS NOMBRE_EMPRESA,"&_
   "(CASE WHEN A.CON_FRANQUICIA=1 and A.CON_OTIC=1 and A.ID_OTIC<>0 THEN "&_ 
   "(select eotic.RUT from EMPRESAS eotic where eotic.ID_EMPRESA=A.ID_OTIC) ELSE 'No Aplica' end) as rut_otic, "&_
   "(CASE WHEN A.CON_FRANQUICIA=1 and A.CON_OTIC=1 and A.ID_OTIC<>0 THEN  "&_
   "(select dbo.MayMinTexto (eotic.R_SOCIAL) from EMPRESAS eotic where eotic.ID_EMPRESA=A.ID_OTIC) "&_
   "ELSE 'No Aplica' end) as nombre_otic,f.MONTO,a.TIPO_DOC,(CASE WHEN A.TIPO_DOC='0' then 'Orden de Compra'"&_
   " WHEN A.TIPO_DOC='1' then 'Vale Vista' "&_
   " WHEN A.TIPO_DOC='2' then 'Depósito Cheque' "&_
   " WHEN A.TIPO_DOC='3' then 'Transferencia' "&_
   " WHEN A.TIPO_DOC='4' then 'Carta Compromiso'  "&_
   " END) as 'Tipo Documento',a.ORDEN_COMPRA,a.VALOR_OC,a.FECHA_REV_CIERRE FROM FACTURAS F "&_
   "  inner join AUTORIZACION A on A.ID_AUTORIZACION=F.ID_AUTORIZACION "&_
   "  inner join EMPRESAS EI ON EI.ID_EMPRESA=A.ID_EMPRESA "&_
   "  inner join EMPRESAS E on E.ID_EMPRESA=F.ID_EMPRESA "&_
   "  inner join PROGRAMA P on P.ID_PROGRAMA=A.ID_PROGRAMA "&_
   "  inner join CURRICULO C on C.ID_MUTUAL=P.ID_MUTUAL  "&_
   "  inner join bloque_programacion BQ on BQ.id_bloque=A.ID_BLOQUE  "&_
   "  WHERE A.ESTADO in (0,1) "&bdoc&txfcd&txfch&txfed&txfeh&txe&txc&_
   " order by f.FACTURA"

	set rs =  conn.execute(qs)%>
    <table border="1">
    <tr>
      <td width="120" align="center"><b><font size="2">RUT</font></b></td>
      <td width="600" align="center"><b><font size="2">RAZON SOCIAL</font></b></td>
      <td width="130" align="center"><b><font size="2">FECHA INICIO</font></b></td>
      <td width="130" align="center"><b><font size="2">FECHA TERMINO</font></b></td>  
      <td width="130" align="center"><b><font size="2">CODIGO</font></b></td>
      <td width="600" align="center"><b><font size="2">NOMBRE CURSO</font></b></td>    
      <td width="100" align="center"><b><font size="2">N&deg; PART</font></b></td>
      <td width="130" align="center"><b><font size="2">FECHA EMISION</font></b></td>           
      <td width="130" align="center"><b><font size="2">FACTURA</font></b></td>         
      <td width="130" align="center"><b><font size="2">ESTADO</font></b></td>  
      <td width="130" align="center"><b><font size="2">DOCUMENTO</font></b></td>     
      <td width="130" align="center"><b><font size="2">MONTO</font></b></td>    
      <td width="250" align="center"><b><font size="2">TIPO DE DOCUMENTO</font></b></td>  
      <td width="250" align="center"><b><font size="2">N&deg; DE DOCUMENTO</font></b></td>     
      <td width="300" align="center"><b><font size="2">FECHA REVISI&Oacute;N Y CIERRE</font></b></td>   
    </tr>
	<%
    while not rs.eof
    %>
        <tr>
          <td align="left"><font size="2"><%=rs("rut")%></font></td>
          <td><font size="2"><%=rs("empresa")%></font></td>
          <td align="center"><font size="2"><%=rs("FECHA_INICIO")%></font></td>          
          <td align="center"><font size="2"><%=rs("FECHA_TERMINO")%></font></td>
          <td><font size="2"><%=rs("CODIGO")%></font></td>    
          <td><font size="2"><%=rs("nombre")%></font></td>  
          <td><font size="2"><%=rs("N_PARTICIPANTES")%></font></td> 
          <td><font size="2"><%=rs("FECHA_EMISION")%></font></td>                 
          <td><font size="2"><%=rs("FACTURA")%></font></td>  
          <td><font size="2"><%=rs("DESCEST")%></font></td>
          <td align="center"><font size="2"><a href="<%=rs("doc")%>">Ver</a></font></td>  
          <td><font size="2"><%=rs("MONTO")%></font></td>    
          <td><font size="2"><%=rs("Tipo Documento")%></font></td>  
          <td><font size="2"><%=rs("ORDEN_COMPRA")%></font></td>          
          <td><font size="2"><%=rs("FECHA_REV_CIERRE")%></font></td>                                          
        </tr>
		<%
	rs.MoveNext
wend
%>
</table>