<%
on error resume next
%>
<!--#include file="../conexion.asp"-->
<%
Response.CodePage = 65001
Response.CharSet = "utf-8"
vautorizacion = Request("id")

dim query
query= "SELECT CURRICULO.NOMBRE_CURSO,CONVERT(VARCHAR(10),PROGRAMA.FECHA_INICIO_, 105) as FECHA_INICIO_, "
query= query&"CONVERT(VARCHAR(10),PROGRAMA.FECHA_TERMINO, 105) as FECHA_TERMINO, "
query= query&"EMPRESAS.R_SOCIAL,EMPRESAS.RUT,PROGRAMA.ID_PROGRAMA,EMPRESAS.ID_EMPRESA,AUTORIZACION.ID_OTIC,"
query= query&"CURRICULO.CODIGO,CURRICULO.SENCE,AUTORIZACION.DOCUMENTO_COMPROMISO as doc,AUTORIZACION.TIPO_DOC, "
query= query&"AUTORIZACION.ORDEN_COMPRA,AUTORIZACION.VALOR_OC,AUTORIZACION.N_PARTICIPANTES,AUTORIZACION.VALOR_CURSO, "
query= query&"dbo.MayMinTexto (INSTRUCTOR_RELATOR.NOMBRES+' '+INSTRUCTOR_RELATOR.A_PATERNO) as instructor,"
query= query&"(CASE WHEN SEDES.ID_SEDE =  27 THEN bloque_programacion.nom_sede "
query= query&" WHEN SEDES.ID_SEDE <>  27 THEN SEDES.NOMBRE+', '+SEDES.DIRECCION+', '+SEDES.CIUDAD END) as sede,"
query= query&"AUTORIZACION.CON_OTIC,bloque_programacion.id_relator,bloque_programacion.id_sede,AUTORIZACION.ID_BLOQUE, "
query= query&"(CASE WHEN getdate()>=PROGRAMA.FECHA_TERMINO then '1' "
query= query&" WHEN getdate()<PROGRAMA.FECHA_TERMINO then '0' END) as 'Validar', "
query= query&"AUTORIZACION.con_franquicia,AUTORIZACION.n_reg_sence"
query= query&" FROM AUTORIZACION " 
query= query&" INNER JOIN PROGRAMA ON PROGRAMA.ID_PROGRAMA=AUTORIZACION.ID_PROGRAMA " 
query= query&" INNER JOIN CURRICULO ON CURRICULO.ID_MUTUAL=PROGRAMA.ID_MUTUAL " 
query= query&" INNER JOIN EMPRESAS ON EMPRESAS.ID_EMPRESA=AUTORIZACION.id_empresa " 
query= query&" inner join bloque_programacion on bloque_programacion.id_bloque=AUTORIZACION.ID_BLOQUE "
query= query&" inner join INSTRUCTOR_RELATOR on INSTRUCTOR_RELATOR.ID_INSTRUCTOR=bloque_programacion.id_relator "
query= query&" inner join SEDES on SEDES.ID_SEDE=bloque_programacion.id_sede "
query= query&" where AUTORIZACION.ID_AUTORIZACION="&vautorizacion

set rsEmp = conn.execute (query)

fecha = year(now)&right("0"&month(now()),2)&right("0"&day(now()),2)
fecha = fecha&right("0"&hour(now()),2)&right("0"&minute(now()),2)&right("0"&second(now()),2)

'conn.execute ("CREATE TABLE #TablaTemporal"&fecha&" (Campo1 int, Campo2 varchar(50))")
'conn.execute ("INSERT INTO #TablaTemporal"&fecha&"  VALUES (1,'Primer campo')")

'set rsEmp1 = conn.execute ("SELECT Campo2 FROM #TablaTemporal20100730164734")

'tab2=rsEmp1("Campo2")

if not rsEmp.eof and not rsEmp.bof then 
%>
 <form name="frminscripcion" id="frminscripcion" action="inscripcionactivas/modInscripcion.asp" method="post">
    <table cellspacing="1" cellpadding="1" border=0>
      <tr>
       <td>Rut Empresa :</td>
       <td><%=replace(FormatNumber(mid(rsEmp("RUT"), 1,len(rsEmp("RUT"))-2),0)&mid(rsEmp("RUT"), len(rsEmp("RUT"))-1,len(rsEmp("RUT"))),",",".")%></td>
       <td><input type="hidden" id="Empresa" name="Empresa" value="<%=rsEmp("ID_EMPRESA")%>"/></td>
       <td>Raz&oacute;n Social : </td>
       <td colspan="2"><%=rsEmp("R_SOCIAL")%></td>
    </tr>
    <tr>
       <td>C&oacute;digo Curso :</td>
       <td><%=rsEmp("CODIGO")%></td>
       <td><input type="hidden" id="Programa" name="Programa" value="<%=rsEmp("ID_PROGRAMA")%>"/></td>
       <td>Nombre Curso :</td>
       <td colspan="2"><%=rsEmp("NOMBRE_CURSO")%></td>
    </tr>
    <tr>
       <td width="100">Fecha Inicio :</td>
       <td width="180"><label id="txtFechI" name="txtFechI"><%=rsEmp("FECHA_INICIO_")%></label></td>
       <td width="20"><input type="hidden" id="autorizacionId" name="autorizacionId" value="<%=vautorizacion%>"/></td>
       <td width="105">Fecha T&eacute;rmino : </td>
       <td width="320"><label id="txtFechT" name="txtFechT"><%=rsEmp("FECHA_TERMINO")%></label></td>
       <td width="190" align="right"><a href="#" onclick="cambiaRelSede(1);">Cambiar Fechas</a></td>
    </tr>
     <tr>
       <td colspan="2">Relator :&nbsp;&nbsp;<label id="txtRelProg" name="txtRelProg"><%=rsEmp("instructor")%></label></td>
       <td><input type="hidden" id="COfran" name="COfran" value="<%=rsEmp("con_franquicia")%>"/></td>
       <td colspan="2">Sede :&nbsp;&nbsp;<label id="txtSalProg" name="txtSalProg"><%=rsEmp("sede")%></label></td>
       <td align="right"><a href="#" onclick="cambiaRelSede(0);">Cambiar Relator/Sede</a></td>
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
       <td width="283">Utiliza Franquicia Sence? :&nbsp;
       	   <input type="radio" name="COfranS" id="COfranS1" value="1" <%=ckdFranSi%> onclick="$('#COfran').val(this.value);MuestraFilas(1);"/>Si
           <input type="radio" name="COfranS" id="COfranS0" value="0" <%=ckdFranNo%> onclick="$('#COfran').val(this.value);MuestraFilas(0);"/>No</td>
       <td width="20"><input type="hidden" id="ConOtic" name="ConOtic" value="<%=rsEmp("con_otic")%>"/></td>
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
	   
	   if(rsEmp("ID_OTIC")<>"0")then %>   
       	    <input type="radio" name="COtic" id="COtic1" value="1" <%=checkedSi%> onclick="$('#ConOtic').val(this.value);$('#Mandante').show();"/>Si
            <input type="radio" name="COtic" id="COtic0" value="0" <%=checkedNo%> onclick="$('#ConOtic').val(this.value);$('#Mandante').hide();"/>No</td>
       <% else %>
       		<input type="radio" name="COtic" id="COtic1" value="1" disabled="disabled"/>Si
            <input type="radio" name="COtic" id="COtic0" value="0" <%=checkedNo%>/>No</td>
       <% end if %>
       <td id="Regpreg" <%=habfilas%>>N&deg; de Registro Sence:&nbsp;&nbsp;<input type="text" id="NReg_Fran" name="NReg_Fran" maxlength="50" size="20" value="<%=rsEmp("N_REG_SENCE")%>"/></td>
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
       <td colspan="6">&nbsp;<input id="tabFechaPart" name="tabFechaPart" type="hidden" value="<%=fecha%>"/>
       <input type="hidden" id="ValInscripcion" name="ValInscripcion" value="<%=rsEmp("Validar")%>"/>
       </td>
     </tr>
   </table>
<table id="list2"></table> 
<div id="pager2"></div> 
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
       <td><label id="txtNParticipantes" name="txtNParticipantes"><%=rsEmp("N_PARTICIPANTES")%></label>
       <input type="hidden" id="NParticipantes" name="NParticipantes" value="<%=rsEmp("N_PARTICIPANTES")%>"/></td>
	   <td>&nbsp;</td>
       <td>Valor por Alumno : </td>
       <td><label id="txtValor" name="txtValor"><%="$ "&replace(FormatNumber(rsEmp("VALOR_CURSO"),0),",",".")%></label>
       <input type="hidden" id="Valor" name="Valor" value="<%=rsEmp("VALOR_CURSO")%>"/></td>
       <td>&nbsp;</td>
       <td>Valor Total :</td>
       <td><label id="txtValorTotal" name="txtValorTotal"><%="$ "&replace(FormatNumber(rsEmp("VALOR_OC"),0),",",".")%></label>
       <input type="hidden" id="ValorTotal" name="ValorTotal" value="<%=rsEmp("VALOR_OC")%>"/></td>
     </tr>
      <tr>
       <td>Tipo Compromiso : </td>
       <td><select id="Compromiso" name="Compromiso" tabindex="1">
           </select><input type="hidden" id="txtTipoDoc" name="txtTipoDoc" value="<%=rsEmp("TIPO_DOC")%>"/></td>
	   <td>&nbsp;</td>
       <td>N&uacute;mero OC/CI/DEP : </td>
       <td><input type="text" id="txtNOrden" name="txtNOrden" value="<%=rsEmp("ORDEN_COMPRA")%>" size="17" tabindex="2"/></td>
       <td><input type="hidden" id="txtNOrdenBuscar" name="txtNOrdenBuscar" value="<%=rsEmp("ORDEN_COMPRA")%>"/></td>    
       <td>Ver Documento :</td>
       <td><a href="#" onclick="documento('<%=rsEmp("doc")%>');">Ver Documento</a></td>
     </tr>
     <tr>
	   <td colspan="6"><input type="hidden" id="txtBloque" name="txtBloque" value="<%=rsEmp("ID_BLOQUE")%>"/>
       <input type="hidden" id="txtSala" name="txtSala" value="<%=rsEmp("id_relator")%>"/>
       <input type="hidden" id="txtRelator" name="txtRelator" value="<%=rsEmp("id_sede")%>"/>
       <input type="hidden" id="txtDocOculto" name="txtDocOculto" value="archivo.pdf"/>
       <input type="hidden" id="DocActual" name="DocActual" value="<%=rsEmp("doc")%>"/>
       </td>
       <td>Cambiar Documento :</td>
       <td><input type="file" name="SubDocSelec" id="SubDocSelec" tabindex="3" maxlength="200" size="18" onchange="DocCambiado(this.value);"/></td>
      </tr>
      <tr>
       <td colspan="8"><center><img id="loading" src="../images/loader.gif" style="display:none;"/></center></td>
     </tr>
   </table>
</form> 
<%
end if
%>
<div id="messageBox2" style="height:60px;overflow:auto;width:450px;"> 
  	<ul></ul> 
</div> 
 