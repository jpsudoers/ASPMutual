<%
on error resume next
%>
<!--#include file="../conexion.asp"-->
<%
Response.CodePage = 65001
Response.CharSet = "utf-8"
vautorizacion = Request("id")

dim query
query= "SELECT CONVERT(VARCHAR(10),PROGRAMA.FECHA_INICIO_, 105) as FECHA_INICIO_, "
query= query&"CONVERT(VARCHAR(10),PROGRAMA.FECHA_TERMINO, 105) as FECHA_TERMINO, "
query= query&"EMPRESAS.R_SOCIAL,EMPRESAS.RUT,AUTORIZACION.ID_OTIC,"
query= query&"CURRICULO.CODIGO,dbo.MayMinTexto (CURRICULO.NOMBRE_CURSO) as NOMBRE_CURSO,CURRICULO.ID_MUTUAL, "
query= query&"AUTORIZACION.DOCUMENTO_COMPROMISO as doc,(CASE WHEN AUTORIZACION.TIPO_DOC='0' then 'Orden de Compra'"
query= query&" WHEN AUTORIZACION.TIPO_DOC='1' then 'Vale Vista' "
query= query&" WHEN AUTORIZACION.TIPO_DOC='2' then 'Depósito Cheque' "
query= query&" WHEN AUTORIZACION.TIPO_DOC='3' then 'Transferencia' "
query= query&" WHEN AUTORIZACION.TIPO_DOC='4' then 'Carta Compromiso' END) as 'Tipo Documento',"
query= query&"AUTORIZACION.ORDEN_COMPRA,AUTORIZACION.VALOR_OC,AUTORIZACION.N_PARTICIPANTES,AUTORIZACION.VALOR_CURSO, "
query= query&"dbo.MayMinTexto (INSTRUCTOR_RELATOR.NOMBRES+' '+INSTRUCTOR_RELATOR.A_PATERNO) as instructor,"
query= query&"(CASE WHEN SEDES.ID_SEDE =  27 THEN bloque_programacion.nom_sede "
query= query&" WHEN SEDES.ID_SEDE <>  27 THEN SEDES.NOMBRE+', '+SEDES.DIRECCION+', '+SEDES.CIUDAD END) as sede,"
query= query&"AUTORIZACION.CON_OTIC,AUTORIZACION.con_franquicia,AUTORIZACION.n_reg_sence,AUTORIZACION.ID_BLOQUE,"
query= query&"EMPRESAS.giro,EMPRESAS.direccion,EMPRESAS.fono,EMPRESAS.email FROM AUTORIZACION " 
query= query&" INNER JOIN PROGRAMA ON PROGRAMA.ID_PROGRAMA=AUTORIZACION.ID_PROGRAMA " 
query= query&" INNER JOIN CURRICULO ON CURRICULO.ID_MUTUAL=PROGRAMA.ID_MUTUAL " 
query= query&" INNER JOIN EMPRESAS ON EMPRESAS.ID_EMPRESA=AUTORIZACION.id_empresa " 
query= query&" inner join bloque_programacion on bloque_programacion.id_bloque=AUTORIZACION.ID_BLOQUE "
query= query&" inner join INSTRUCTOR_RELATOR on INSTRUCTOR_RELATOR.ID_INSTRUCTOR=bloque_programacion.id_relator "
query= query&" inner join SEDES on SEDES.ID_SEDE=bloque_programacion.id_sede "
query= query&" where AUTORIZACION.ID_AUTORIZACION="&vautorizacion

set rsEmp = conn.execute (query)

if not rsEmp.eof and not rsEmp.bof then 
%>
 <form name="frminscripcion" id="frminscripcion" action="#" method="post">
    <table cellspacing="1" cellpadding="1" border=0>
     <tr>
       <td>Rut Empresa :</td>
       <td><%=replace(FormatNumber(mid(rsEmp("RUT"), 1,len(rsEmp("RUT"))-2),0)&mid(rsEmp("RUT"), len(rsEmp("RUT"))-1,len(rsEmp("RUT"))),",",".")%></td>
       <td><input type="hidden" id="id_autorizacion" name="id_autorizacion" value="<%=vautorizacion%>"/></td>
       <td>Raz&oacute;n Social : </td>
       <td><%=rsEmp("R_SOCIAL")%></td>
    </tr>
    <tr>
       <td>Giro :</td>
       <td><%=rsEmp("giro")%></td>
       <td>&nbsp;</td>
       <td>Direcci&oacute;n :</td>
       <td><%=rsEmp("direccion")%></td>
    </tr>
    <tr>
       <td width="100">Tel&eacute;fono :</td>
       <td width="180"><%=rsEmp("Fono")%></td>
       <td width="20">&nbsp;</td>
       <td width="105">Email : </td>
       <td width="495"><%=rsEmp("email")%></td>
    </tr>
    <tr>
       <td colspan="5">&nbsp;</td>
    </tr>    
   </table>
   
   <div style="width:900px;"> 
      <table id="mytable" border="1" width="900">
      <thead> 
      <tr>
           <th>Fecha</th> 
           <th>Rut</th> 
           <th>Empresa / OTIC</th> 
           <th>N&deg; Factura</th> 
           <th>Estado</th>            
      </tr>
      </thead> 
      <tbody> 
           <%
		   
				sql = "SELECT E.RUT as rut,dbo.MayMinTexto (E.R_SOCIAL) as empresa, "
				sql = sql&"CONVERT(VARCHAR(10),P.FECHA_INICIO_, 105) as FECHA_INICIO,CONVERT(VARCHAR(10),F.FECHA_EMISION,105) as FECHA," 
				sql = sql&"F.FECHA_EMISION,C.CODIGO, dbo.MayMinTexto (C.NOMBRE_CURSO) as nombre, A.N_PARTICIPANTES, "
				sql = sql&"A.ID_AUTORIZACION,A.ID_PROGRAMA as ID_PROGRAMA,A.DOCUMENTO_COMPROMISO as doc, F.FACTURA,F.ID_FACTURA,"
				sql = sql&"F.ESTADO,(CASE F.ESTADO WHEN 1 THEN 'Vigente' ELSE 'Anulada' end) as 'DESCEST',"
				sql = sql&" EI.RUT AS RUT_EMPRESA,EI.R_SOCIAL AS NOMBRE_EMPRESA,"
				sql = sql&"(CASE WHEN A.CON_FRANQUICIA=1 and A.CON_OTIC=1 and A.ID_OTIC<>0 THEN " 
				sql = sql&"(select eotic.RUT from EMPRESAS eotic where eotic.ID_EMPRESA=A.ID_OTIC) ELSE 'No Aplica' end) as rut_otic, "
				sql = sql&"(CASE WHEN A.CON_FRANQUICIA=1 and A.CON_OTIC=1 and A.ID_OTIC<>0 THEN " 
				sql = sql&"(select dbo.MayMinTexto (eotic.R_SOCIAL) from EMPRESAS eotic where eotic.ID_EMPRESA=A.ID_OTIC) "
				sql = sql&" ELSE 'No Aplica' end) as nombre_otic FROM FACTURAS F "
				sql = sql&" inner join AUTORIZACION A on A.ID_AUTORIZACION=F.ID_AUTORIZACION "
				sql = sql&" inner join EMPRESAS EI ON EI.ID_EMPRESA=A.ID_EMPRESA "
				sql = sql&" inner join EMPRESAS E on E.ID_EMPRESA=F.ID_EMPRESA "
				sql = sql&" inner join PROGRAMA P on P.ID_PROGRAMA=A.ID_PROGRAMA "  
				sql = sql&" inner join CURRICULO C on C.ID_MUTUAL=P.ID_MUTUAL " 
				sql = sql&" inner join bloque_programacion BQ on BQ.id_bloque=A.ID_BLOQUE " 
				sql = sql&" WHERE A.ESTADO=0 AND F.ID_AUTORIZACION="&vautorizacion			   
			   
			   set rsSol =  conn.execute(sql)
		   while not rsSol.eof
	      %>
          <tr>
           <td><%=rsSol("FECHA")%></td> 
           <td><%=rsSol("rut")%></td> 
		   <td><%=rsSol("empresa")%></td> 
           <td><%=rsSol("FACTURA")%></td> 
		   <td><%=rsSol("DESCEST")%></td>            
          </tr>
          <%
		  	   rsSol.Movenext
			wend
		  %>
       </tbody> 
     </table>
   </div>    
</form> 
<%
end if
%> 