<!--#include file="../conexion.asp"-->
<%
fecha = year(now)&right("0"&month(now()),2)&right("0"&day(now()),2)
fecha = fecha&right("0"&hour(now()),2)&right("0"&minute(now()),2)&right("0"&second(now()),2)
Response.ContentType ="application/vnd.ms-excel"
Response.AddHeader "content-disposition", "inline; filename=FACT_"&fecha&".xls"
	

sql = "SELECT E.RUT as rut,dbo.MayMinTexto (E.R_SOCIAL) as empresa, "
sql = sql&"CONVERT(VARCHAR(10),P.FECHA_INICIO_, 105) as FECHA_INICIO,CONVERT(VARCHAR(10),F.FECHA_EMISION,105) as FECHA," 
sql = sql&"(CASE WHEN A.TIPO_DOC='0' then 'Orden de Compra'"
sql = sql&" WHEN A.TIPO_DOC='1' then 'Vale Vista' "
sql = sql&" WHEN A.TIPO_DOC='2' then 'Depósito Cheque' "
sql = sql&" WHEN A.TIPO_DOC='3' then 'Transferencia' "
sql = sql&" WHEN A.TIPO_DOC='4' then 'Carta Compromiso' " 
sql = sql&" END) as 'Tipo Documento',A.ORDEN_COMPRA,"
sql = sql&"F.FECHA_EMISION,C.CODIGO, dbo.MayMinTexto (C.NOMBRE_CURSO) as nombre, A.N_PARTICIPANTES, "
sql = sql&"A.ID_AUTORIZACION,A.ID_PROGRAMA as ID_PROGRAMA,A.DOCUMENTO_COMPROMISO as doc, F.FACTURA,F.ID_FACTURA,"
sql = sql&"F.ESTADO,(CASE F.ESTADO WHEN 1 THEN 'Vigente' WHEN 0 THEN 'Anulada' WHEN 2 THEN 'Vig/Ref' end) as 'DESCEST',"
sql = sql&" EI.RUT AS RUT_EMPRESA,EI.R_SOCIAL AS NOMBRE_EMPRESA,"
sql = sql&"(CASE WHEN A.CON_FRANQUICIA=1 and A.CON_OTIC=1 and A.ID_OTIC<>0 THEN " 
sql = sql&"(select eotic.RUT from EMPRESAS eotic where eotic.ID_EMPRESA=A.ID_OTIC) ELSE 'No Aplica' end) as rut_otic, "
sql = sql&"(CASE WHEN A.CON_FRANQUICIA=1 and A.CON_OTIC=1 and A.ID_OTIC<>0 THEN " 
sql = sql&"(select dbo.MayMinTexto (eotic.R_SOCIAL) from EMPRESAS eotic where eotic.ID_EMPRESA=A.ID_OTIC) "
sql = sql&" ELSE 'No Aplica' end) as nombre_otic,E.TIPO as tipo_emp,N_TICKETS,F.MONTO,P.FECHA_INICIO_,P.FECHA_TERMINO,e.tipo,a.n_reg_sence FROM FACTURAS F "
sql = sql&" inner join AUTORIZACION A on A.ID_AUTORIZACION=F.ID_AUTORIZACION "
sql = sql&" inner join EMPRESAS EI ON EI.ID_EMPRESA=A.ID_EMPRESA "
sql = sql&" inner join EMPRESAS E on E.ID_EMPRESA=F.ID_EMPRESA "
sql = sql&" inner join PROGRAMA P on P.ID_PROGRAMA=A.ID_PROGRAMA "  
sql = sql&" inner join CURRICULO C on C.ID_MUTUAL=P.ID_MUTUAL " 
sql = sql&" inner join bloque_programacion BQ on BQ.id_bloque=A.ID_BLOQUE " 
sql = sql&" WHERE /*A.ESTADO=0 and */Year(F.FECHA_EMISION)=2015 and Month(F.FECHA_EMISION)=5 "
sql = sql&" order by F.FECHA_EMISION asc"

	set rsHc =  conn.execute(sql)%>

    <table width="4000" border="1">
	<%
    while not rsHc.eof
    %>
    <tr>
      <td align="center"><b><font size="2"><%=rsHc("ID_FACTURA")%></font></b></td>
      <td align="center"><b><font size="2"><%=rsHc("FECHA")%></font></b></td>     
      <td align="right"><b><font size="2"><%=rsHc("rut")%></font></b></td>
      <td><b><font size="2"><%=rsHc("empresa")%></font></b></td>
      <td><b><font size="2"><%=rsHc("CODIGO")%></font></b></td>
      <td><b><font size="2"><%=rsHc("nombre")%></font></b></td>
      <td><b><font size="2"><%=rsHc("FACTURA")%></font></b></td>
      <td><b><font size="2"><%=rsHc("DESCEST")%></font></b></td> 
      <td><b><font size="2"><%=rsHc("MONTO")%></font></b></td>
      <td align="center"><b><font size="2"><%=rsHc("FECHA_INICIO_")%></font></b></td>
      <td align="center"><b><font size="2"><%=rsHc("FECHA_TERMINO")%></font></b></td>
      <td align="center"><b><font size="2"><%=rsHc("tipo")%></font></b></td> 
      <td align="center"><b><font size="2"><%=rsHc("Tipo Documento")%></font></b></td>
      <td align="center"><b><font size="2"><%=rsHc("ORDEN_COMPRA")%></font></b></td>    
      <td align="center"><b><font size="2"><%=rsHc("n_reg_sence")%></font></b></td>           
    </tr>	
    <%
        rsHc.MoveNext
    wend
    %>
    </table>
