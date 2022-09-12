<!--#include file="../conexion.asp"-->
<%
fecha = year(now)&right("0"&month(now()),2)&right("0"&day(now()),2)
fecha = fecha&right("0"&hour(now()),2)&right("0"&minute(now()),2)&right("0"&second(now()),2)
Response.ContentType ="application/vnd.ms-excel"
Response.AddHeader "content-disposition", "inline; filename=DET_FACTURACION_"&fecha&".xls"

	vid = Request("id_prog")
	vrelator = Request("id_relator")

	dim sqlOtic
	dim facOtic
	dim facEmp
	dim sql
	sql="select EMPRESAS.RUT,UPPER (EMPRESAS.R_SOCIAL) as R_SOCIAL,AUTORIZACION.CON_OTIC,AUTORIZACION.ID_OTIC, "
	sql = sql&"AUTORIZACION.ID_EMPRESA,AUTORIZACION.ID_AUTORIZACION from AUTORIZACION "  
	sql = sql&" inner join EMPRESAS on EMPRESAS.ID_EMPRESA=AUTORIZACION.ID_EMPRESA "  
	sql = sql&" inner join bloque_programacion on bloque_programacion.id_bloque=AUTORIZACION.ID_BLOQUE "  
	sql = sql&" where bloque_programacion.id_bloque=(select BP.id_bloque from bloque_programacion BP "
	sql = sql&" where BP.id_programa="&vid&" AND BP.id_relator="&vrelator&")"
	sql = sql&" ORDER BY EMPRESAS.R_SOCIAL asc"
	
	set rsEmp =  conn.execute(sql)
		%>
        <table width="930" cellspacing="0" cellpadding="2" border="1">
        <tr>
          <td width="90" align="center"><b>Rut</b></td>
          <td width="390" align="center"><b>Raz&oacute;n Social</b></td>
          <td width="80" align="center"><b>Monto</b></td>
          <td width="100" align="center"><b>N° de Factura</b></td>
          <td width="120" align="center"><b>Fecha de Emisi&oacute;n</b></td>
          <td width="150" align="center"><b>Fecha de Vencimiento</b></td>
        </tr>
        <%   
		while not rsEmp.eof
		%>
		<tr>
		  <td align="center"><b><font size="2"><%=replace(FormatNumber(mid(rsEmp("rut"), 1,len(rsEmp("rut"))-2),0)&mid(rsEmp("rut"), len(rsEmp("rut"))-1,len(rsEmp("rut"))),",",".")%></font></b></td>
		  <td><b><font size="2"><%=rsEmp("R_SOCIAL")%></font></b></td>
            <%  
			facEmp="select CONVERT(VARCHAR(10),F.FECHA_EMISION, 105) as FECHA_EMISION,"
			facEmp=facEmp&"CONVERT(VARCHAR(10),F.FECHA_VENCIMIENTO, 105) as FECHA_VENCIMIENTO,F.MONTO,F.FACTURA from FACTURAS F "
			facEmp=facEmp&"where F.ID_EMPRESA='"&rsEmp("ID_EMPRESA")&"' and F.ID_AUTORIZACION='"&rsEmp("ID_AUTORIZACION")&"'"
			
			set rsFacEmp =  conn.execute(facEmp)
			
			if not rsFacEmp.eof and not rsFacEmp.bof then%>
			  <td><b><font size="2"><%="$"&replace(FormatNumber(rsFacEmp("MONTO"),0),",",".")%></font></b></td>
			  <td align="center"><b><font size="2"><%=rsFacEmp("FACTURA")%></font></b></td>
			  <td align="center"><b><font size="2"><%=rsFacEmp("FECHA_EMISION")%></font></b></td>
			  <td align="center"><b><font size="2"><%=rsFacEmp("FECHA_VENCIMIENTO")%></font></b></td>
			</tr>	
			<%else%>
              <td align="center"><b><font size="2">-</font></b></td>
              <td align="center"><b><font size="2">Por Facturar</font></b></td>
              <td align="center"><b><font size="2">-</font></b></td>
              <td align="center"><b><font size="2">-</font></b></td>
            <%end if%>
		</tr>	
        <%  
		if(rsEmp("CON_OTIC")="1")then
			sqlOtic="select E.RUT,UPPER (E.R_SOCIAL) as R_SOCIAL from EMPRESAS E where E.ID_EMPRESA="&rsEmp("ID_OTIC")
		
			set rsOTIC =  conn.execute(sqlOtic)
		%>
		<tr>
		  <td align="center"><b><font size="2"><%=replace(FormatNumber(mid(rsOTIC("rut"), 1,len(rsOTIC("rut"))-2),0)&mid(rsOTIC("rut"), len(rsOTIC("rut"))-1,len(rsOTIC("rut"))),",",".")%></font></b></td>
		  <td><b><font size="2"><%=rsOTIC("R_SOCIAL")%></font></b></td>
            <%  
			facOtic="select CONVERT(VARCHAR(10),F.FECHA_EMISION, 105) as FECHA_EMISION,"
			facOtic=facOtic&"CONVERT(VARCHAR(10),F.FECHA_VENCIMIENTO, 105) as FECHA_VENCIMIENTO,F.MONTO,F.FACTURA from FACTURAS F "
			facOtic=facOtic&"where F.ID_EMPRESA='"&rsEmp("ID_OTIC")&"' and F.ID_AUTORIZACION='"&rsEmp("ID_AUTORIZACION")&"'"
			
			set rsFacOtic =  conn.execute(facOtic)
			
			if not rsFacOtic.eof and not rsFacOtic.bof then%>
			  <td><b><font size="2"><%="$"&replace(FormatNumber(rsFacOtic("MONTO"),0),",",".")%></font></b></td>
			  <td align="center"><b><font size="2"><%=rsFacOtic("FACTURA")%></font></b></td>
			  <td align="center"><b><font size="2"><%=rsFacOtic("FECHA_EMISION")%></font></b></td>
			  <td align="center"><b><font size="2"><%=rsFacOtic("FECHA_VENCIMIENTO")%></font></b></td>
			</tr>	
			<%else%>
              <td align="center"><b><font size="2">-</font></b></td>
              <td align="center"><b><font size="2">Por Facturar</font></b></td>
              <td align="center"><b><font size="2">-</font></b></td>
              <td align="center"><b><font size="2">-</font></b></td>
            <%end if
		end if
			rsEmp.MoveNext
		wend
		%>
        </table>   