<!--#include file="../conexion.asp"-->
<%
fecha = year(now)&right("0"&month(now()),2)&right("0"&day(now()),2)
fecha = fecha&right("0"&hour(now()),2)&right("0"&minute(now()),2)&right("0"&second(now()),2)
Response.ContentType ="application/vnd.ms-excel"
Response.AddHeader "content-disposition", "inline; filename=Carga_Netsuite_"&fecha&".xls"
	
qsHc=" SELECT CONVERT(VARCHAR(10),F.FECHA_EMISION,105) AS FECHA_FACTURA,"&_
	 "e.RUT,UPPER(E.R_SOCIAL) as 'Empresa',UPPER(E.DIRECCION) AS DIRECCION,"&_
	 "UPPER(E.DIRECCION+', '+E.CIUDAD) AS DIR, E.ID_EMPRESA,"&_
	 "UPPER(E.COMUNA) AS COMUNA, DBO.MayMinTexto(E.CIUDAD) AS CIUDAD,NOC=dbo.RemoveCharacteres(ORDEN_COMPRA),"&_
	 "E.FONO, UPPER(E.GIRO) AS GIRO,(CASE WHEN a.TIPO_DOC='0' then 'O/C '+a.ORDEN_COMPRA "&_
	 " WHEN a.TIPO_DOC='1' then 'Vale Vista N° '+a.ORDEN_COMPRA "&_
	 " WHEN a.TIPO_DOC='2' then 'DV '+a.ORDEN_COMPRA "&_
	 " WHEN a.TIPO_DOC='3' then 'Transferencia N° '+a.ORDEN_COMPRA "&_
	 " WHEN a.TIPO_DOC='4' then 'CONTRA FACTURA '+a.ORDEN_COMPRA "&_
	 " END) as 'Tipo Documento',f.MONTO,f.VALOR_CURSO,"&_
	 "CONVERT(VARCHAR(10),P.FECHA_INICIO_,105) AS FECHA_INICIO,"&_
	 "CONVERT(VARCHAR(10),P.FECHA_TERMINO,105) AS FECHA_TERMINO, "&_
	 "CONVERT(VARCHAR(10),GETDATE(),105) AS hoy, "&_	 
	 "C.HORAS,C.DESCRIPCION,/*C.*/COD_SOFTLAND=C.[COD_NETSUITE],c.CODIGO,"&_
         "ISNULL(convert(varchar,C.ID_CECO),'') as ID_CECO, "&_ 	 
	 "ISNULL(convert(varchar,A.N_REG_SENCE),'') as N_REG_SENCE, "&_ 
	 "(select (CASE WHEN bq.ID_SEDE =  27 THEN REPLACE(bq.nom_sede,'#','') "&_
	 "WHEN bq.ID_SEDE <>  27 THEN S.DIRECCION+', '+S.CIUDAD END) "&_
	 "from bloque_programacion bq "&_
	 " inner join SEDES s on s.ID_SEDE=bq.id_sede "&_
	 " inner join PROGRAMA p on p.ID_PROGRAMA=bq.id_programa "&_
	 " where bq.id_bloque=A.ID_BLOQUE) as SEDE, "&_
	 " ASIS=(SELECT COUNT(*) FROM HISTORICO_CURSOS HC  "&_
	 " WHERE HC.ID_AUTORIZACION=A.ID_AUTORIZACION AND HC.ASISTENCIA<>0), "&_
	 " INASIS=(SELECT COUNT(*) FROM HISTORICO_CURSOS HC  "&_
	 " WHERE HC.ID_AUTORIZACION=A.ID_AUTORIZACION AND HC.ASISTENCIA=0),F.ID_FACTURA, "&_
	 " ES_OTIC=(case when E.ID_OTIC is null then 1 else 0 end), "&_
	 " NombBenef=(select o.R_SOCIAL from EMPRESAS o where o.ID_EMPRESA=a.ID_EMPRESA),e.ID_INTERNO_NETSUITE, "&_
	 " RutBenef=(select o.RUT from EMPRESAS o where o.ID_EMPRESA=a.ID_EMPRESA),U.ID_USUARIO_NETSUITE,'PAGO'=(CASE WHEN a.TIPO_DOC='0' then '2: Crédito' else '1: Contado' END), "&_
	 " 'Referencia'=ISNULL((case when a.ID_TIPO_REFERENCIA is not null then tr.TIPO_REFERENCIA+' '+a.N_REFERENCIA else '' END),''),"&_
	 " Descripcion2=(select (CASE WHEN bq.ID_SEDE =  27 THEN REPLACE(bq.nom_sede,'#','') "&_
	 " WHEN bq.ID_SEDE <>  27 THEN S.DIRECCION+', '+S.CIUDAD END) "&_
	 " from bloque_programacion bq "&_
	 " inner join SEDES s on s.ID_SEDE=bq.id_sede "&_
	 " inner join PROGRAMA p on p.ID_PROGRAMA=bq.id_programa "&_
	 " where bq.id_bloque=A.ID_BLOQUE)+ char(13)+CHAR(10) + c.CODIGO,A.N_PARTICIPANTES,cc.id_netsuite, "&_
         " B.NOMBRE_BANCO, CC.NUMERO_CUENTA,A.MONTO_TRANSFERENCIA, CONVERT(VARCHAR(10),A.FECHA_TRANSFERENCIA, 105) as FECHA_TRANSFERENCIA,  "&_
         " 'RUT_DEPO'=replace(A.RUT_DEPOSITANTE,'.',''), A.NOMBRE_DEPOSITANTE, A.EMAIL_DEPOSITANTE, A.RELACION_DEPOSITANTE, "&_
	 " contacto=UPPER(e.NOMBRES),contactoEmail=LOWER(e.EMAIL),contactoCargo=UPPER(e.CARGO),contactoFono=e.FONO_CONTACTO, "&_
	 " cod=CONVERT(varchar(30),cc.NUMERO_CUENTA)+CONVERT(varchar(30),a.MONTO_TRANSFERENCIA)+"&_
	 " CONVERT(varchar(10),a.FECHA_TRANSFERENCIA,105)+replace(A.RUT_DEPOSITANTE,'.','')+CONVERT(varchar(30),f.ID_FACTURA)  "&_
	 " FROM FACTURAS F "&_
	 " inner join AUTORIZACION a on a.ID_AUTORIZACION=f.ID_AUTORIZACION "&_
	 "  left join TIPO_REFERENCIA tr on tr.ID_TIPO_REFERENCIA=a.ID_TIPO_REFERENCIA "&_
	 " inner join empresas e on e.ID_EMPRESA=f.ID_EMPRESA "&_
	 " inner join programa p on p.ID_PROGRAMA=a.ID_PROGRAMA "&_
	 " inner join CURRICULO c on c.ID_MUTUAL=p.ID_MUTUAL "&_
	 " LEFT join USUARIOS U on U.ID_USUARIO=F.FACTURADO "&_
	 " left join BANCOS B on B.ID_BANCO=A.ID_BANCO "&_
	 " left join CUENTAS_CORRIENTES CC on CC.ID_BANCO=A.ID_BANCO AND CC.ID_CUENTA_CORRIENTE=A.ID_CUENTA_CORRIENTE "&_
	 " WHERE F.FECHA_EMISION>=CONVERT(date, '01-06-2022') and c.[COD_NETSUITE] is not null /*and e.ID_INTERNO_NETSUITE is not null*/ "&_
	 " order by F.FECHA_EMISION asc"

	set rsHc =  conn.execute(qsHc)

'response.write(qsHc)
'response.end
%>
    <table width="4500" border="1">
    <tr>
      <td width="100" align="center"><b><font size="2">ID Externo</font></b></td>
      <td width="100" align="center"><b><font size="2">ID Externo Emi F (netsuite)</font></b></td>
      <td width="100" align="center"><b><font size="2">ID Externo Pago F (netsuite)</font></b></td>
      <td width="100" align="center"><b><font size="2">FECHA</font></b></td>
      <td width="200" align="center"><b><font size="2">N.º DE OC</font></b></td>
      <td width="100" align="center"><b><font size="2">HES /MIGO</font></b></td>
      <td width="100" align="center"><b><font size="2">ID INTERNO</font></b></td>
      <td width="100" align="center"><b><font size="2">ID REPRESENTANTE VENTA</font></b></td>
      <td width="350" align="center"><b><font size="2">DESCRIPCION DE LA LINEA (REVISAR)</font></b></td>      
      <td width="100" align="center"><b><font size="2">T&Eacute;RMINOS</font></b></td>   
      <td width="100" align="center"><b><font size="2">ID INTERNO - CDR</font></b></td> 
      <td width="100" align="center"><b><font size="2">ID INTERNO - ARTICULOS</font></b></td> 
      <td width="100" align="center"><b><font size="2">TASA</font></b></td> 
      <td width="100" align="center"><b><font size="2">PRECIO UNITARIO</font></b></td> 
      <td width="100" align="center"><b><font size="2">CANTIDAD</font></b></td> 
      <td width="100" align="center"><b><font size="2">ESTADO APROBACI&Oacute;N</font></b></td> 
      <td width="100" align="center"><b><font size="2">ID CUENTA CORRIENTE NETSUITE</font></b></td> 

      <td width="100" align="center"><b><font size="2">ID CLIENTE INTERNO</font></b></td> 
      <td width="100" align="center"><b><font size="2">RUT</font></b></td> 
      <td width="100" align="center"><b><font size="2">RAZÓN SOCIAL</font></b></td> 
      <td width="100" align="center"><b><font size="2">GIRO</font></b></td> 
      <td width="100" align="center"><b><font size="2">DIRECCIÓN</font></b></td> 
      <td width="100" align="center"><b><font size="2">COMUNA</font></b></td> 
      <td width="100" align="center"><b><font size="2">NOMBRE_CONTACTO</font></b></td> 
      <td width="100" align="center"><b><font size="2">EMAIL_CONTACTO</font></b></td> 
      <td width="100" align="center"><b><font size="2">TELEFONO_CONTACTO</font></b></td> 
      <td width="100" align="center"><b><font size="2">CARGO_CONTACTO</font></b></td> 

	  <td width="150" align="center"><b><font size="2">Banco</font></b></td>  
	  <td width="150" align="center"><b><font size="2">N. Cuenta</font></b></td>  
	  <td width="150" align="center"><b><font size="2">Monto Transferencia</font></b></td>  
	  <td width="150" align="center"><b><font size="2">Fecha Transferencia</font></b></td>   
	  <td width="150" align="center"><b><font size="2">Rut Depositante</font></b></td>  
	  <td width="150" align="center"><b><font size="2">Nombre Depositante</font></b></td>  
	  <td width="150" align="center"><b><font size="2">Email Depositante</font></b></td>  
	  <td width="150" align="center"><b><font size="2">Relaci&oacute;n Depositante</font></b></td> 

<td width="150" align="center"><b><font size="2">C&oacute;digo Vaidaci&oacute;n</font></b></td> 
    </tr>
    <%
    while not rsHc.eof


    %>
    <tr>
      <td><%=rsHc("ID_FACTURA")%></td>
      <td><%="F"&rsHc("ID_FACTURA")%></td>
      <td><%="PF"&rsHc("ID_FACTURA")%></td>
      <td><%=rsHc("FECHA_FACTURA")%></td>     
      <td><%=rsHc("NOC")%></td>
      <td><%=rsHc("Referencia")%></td>
      <td><%=rsHc("ID_INTERNO_NETSUITE")%></td>
      <td><%=rsHc("ID_USUARIO_NETSUITE")%></td>
      <td><%="1"& vbCrLf   & "\r\n 2"%></td> 
      <td><%=rsHc("PAGO")%></td>      
      <td>1017</td> 
      <td><%=rsHc("COD_SOFTLAND")%></td>  
      <td><%=rsHc("MONTO")%></td>
      <td><%=rsHc("VALOR_CURSO")%></td>
      <td><%=rsHc("N_PARTICIPANTES")%></td>
      <td>Aprobado</td>
      <td><%=rsHc("id_netsuite")%></td>
      <td><%=rsHc("RUT")&"C"%></td> 
      <td><%=rsHc("RUT")%></td> 
      <td><%=rsHc("Empresa")%></td> 
      <td><%=rsHc("GIRO")%></td> 
      <td><%=rsHc("DIRECCION")%></td> 
      <td><%=rsHc("COMUNA")%></td> 
      <td><%=rsHc("contacto")%></td> 
      <td><%=rsHc("contactoEmail")%></td>  
      <td><%=rsHc("contactoFono")%></td> 
      <td><%=rsHc("contactoCargo")%></td>
      <%
	if(rsHc("PAGO")="1: Contado")then
      %> 
	  <td><%=rsHc("NOMBRE_BANCO")%></td>
	  <td><%=rsHc("NUMERO_CUENTA")%></td>
	  <td><%=rsHc("MONTO_TRANSFERENCIA")%></td>
	  <td><%=rsHc("FECHA_TRANSFERENCIA")%></td>
	  <td><%=rsHc("RUT_DEPO")%></td>
	  <td><%=rsHc("NOMBRE_DEPOSITANTE")%></td>
	  <td><%=rsHc("EMAIL_DEPOSITANTE")%></td>
	  <td><%=rsHc("RELACION_DEPOSITANTE")%></td>
	  <td><%=rsHc("cod")%></td>
	<%else%> 
	  <td></td>
	  <td></td>
	  <td></td>
	  <td></td>
	  <td></td>
	  <td></td>
	  <td></td>
	  <td></td>
	  <td></td>
	<%end if%> 

    </tr>	
    <%
        rsHc.MoveNext
    wend
    %>
    </table>
