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
query= query&"AUTORIZACION.CON_OTIC,AUTORIZACION.con_franquicia,AUTORIZACION.n_reg_sence,AUTORIZACION.ID_BLOQUE"
query= query&" FROM AUTORIZACION " 
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
       <td>C&oacute;digo Curso :</td>
       <td><%=rsEmp("CODIGO")%></td>
       <td>&nbsp;</td>
       <td>Nombre Curso :</td>
       <td><%=rsEmp("NOMBRE_CURSO")%></td>
    </tr>
    <tr>
       <td width="100">Fecha Inicio :</td>
       <td width="180"><%=rsEmp("FECHA_INICIO_")%></td>
       <td width="20">&nbsp;</td>
       <td width="105">Fecha T&eacute;rmino : </td>
       <td width="495"><%=rsEmp("FECHA_TERMINO")%></td>
    </tr>
    <tr>
       <td colspan="2">Relator :&nbsp;&nbsp;<%=rsEmp("instructor")%></td>
       <td>&nbsp;</td>
       <td colspan="2">Sede :&nbsp;&nbsp;<%=rsEmp("sede")%></td>
    </tr>
    </table>
     <table cellspacing="1" cellpadding="1" border=0>
     <tr>
       <%
	   ckdFranSi=""
	   ckdFranNo=""
	   habfilas=""
   
	   if(rsEmp("con_franquicia")="1")then
		  ckdFranSi="checked='checked'"
		  habfilas=""
	   else
		  ckdFranNo="checked='checked'"
		  habfilas="style='display:none'"
	   end if
	   %> 
       <td width="278">Utiliza Franquicia Sence? :&nbsp;
       	   <input type="radio" name="COfranS" id="COfranS1" value="1" <%=ckdFranSi%> disabled="disabled"/>Si
           <input type="radio" name="COfranS" id="COfranS0" value="0" <%=ckdFranNo%> disabled="disabled"/>No</td>
       <td width="20">&nbsp;</td>
       <td width="310" id="Oticpreg" <%=habfilas%>>Con OTIC :&nbsp;&nbsp;
       <%
	   checkedSi=""
	   checkedNo=""
       habFilaOtic=""
	   
	   if(rsEmp("con_otic")="1")then
		   checkedSi="checked='checked'"
	       habFilaOtic=""		   
	   else
		   checkedNo="checked='checked'"
	       habFilaOtic="style='display:none'"				   
	   end if
	   %>   
       <input type="radio" name="COtic" id="COtic1" value="1" <%=checkedSi%> disabled="disabled"/>Si
       <input type="radio" name="COtic" id="COtic0" value="0" <%=checkedNo%> disabled="disabled"/>No</td>
       <td id="Regpreg" <%=habfilas%>>N&deg; de Registro Sence:&nbsp;&nbsp;<%=rsEmp("n_reg_sence")%></td>
     </tr> 
     <%
      if(rsEmp("ID_OTIC")<>"0")then
		 set rsOtic = conn.execute ("select RUT,R_SOCIAL from EMPRESAS where EMPRESAS.ID_EMPRESA="&rsEmp("ID_OTIC"))
	 %>
     <tr name="Mandante" id="Mandante" <%=habFilaOtic%>>
         <td>Rut OTIC :&nbsp;&nbsp;<%=replace(FormatNumber(mid(rsOtic("RUT"), 1,len(rsOtic("RUT"))-2),0)&mid(rsOtic("RUT"), len(rsOtic("RUT"))-1,len(rsOtic("RUT"))),",",".")%></td>
         <td>&nbsp;</td>
         <td colspan="3">Raz&oacute;n Social :&nbsp;&nbsp;<%=rsOtic("R_SOCIAL")%></td>
     </tr>
     <%end if%>
     <tr>
       <td colspan="6">&nbsp;</td>
     </tr>
   </table>
   <table id="listPart"></table> 
   <div id="pagePart"></div> 
   <table cellspacing="0" cellpadding="1" border=0>
     <tr>
	   <td width="150">&nbsp;</td>
       <td width="130"></td>
       <td width="20"></td>
       <td width="140"></td>
       <td width="100"></td>
       <td width="20"></td>
       <td width="150"></td>
       <td width="130"></td>
      </tr>
    <tr>
       <td>N&#176; de Participantes : </td>
       <td><%=rsEmp("N_PARTICIPANTES")%></td>
	   <td>&nbsp;</td>
       <td>Valor por Alumno : </td>
       <td><%="$ "&replace(FormatNumber(rsEmp("VALOR_CURSO"),0),",",".")%></td>
       <td>&nbsp;</td>
       <td>Valor Total :</td>
       <td><%="$ "&replace(FormatNumber(rsEmp("VALOR_OC"),0),",",".")%></td>
     </tr>
      <tr>
       <td>Tipo Compromiso : </td>
       <td><%=rsEmp("Tipo Documento")%></td>
	   <td>&nbsp;</td>
       <td>N&uacute;mero OC/CI/DEP : </td>
       <td><%=rsEmp("ORDEN_COMPRA")%></td>
       <td></td>    
       <td>Ver Documento :</td>
       <td><a href="#" onclick="docOC('<%=rsEmp("doc")%>');">Ver Documento</a></td>
     </tr>
     <tr>
	   <td colspan="8">&nbsp;</td>
      </tr>
   </table>
</form> 
<%
end if
%> 