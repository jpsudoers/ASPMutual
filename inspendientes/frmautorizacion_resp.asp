<%
on error resume next
%>
<!--#include file="../conexion.asp"-->
<%
Response.CodePage = 65001
Response.CharSet = "utf-8"
vid = Request("id")

dim query
query= "SELECT CURRICULO.NOMBRE_CURSO,CONVERT(VARCHAR(10),PROGRAMA.FECHA_INICIO_, 105) as FECHA_INICIO_, "
query= query&"CONVERT(VARCHAR(10),PROGRAMA.FECHA_TERMINO, 105) as FECHA_TERMINO, "
query= query&"EMPRESAS.R_SOCIAL,EMPRESAS.RUT,preinscripciones.id_otic, PROGRAMA.ID_PROGRAMA,EMPRESAS.ID_EMPRESA,"
query= query&"CURRICULO.CODIGO,CURRICULO.SENCE,preinscripciones.tipo_compromiso,preinscripciones.numero_compromiso, "
query= query&"preinscripciones.id_preinscripcion,preinscripciones.doc_compromiso as doc,preinscripciones.participantes,"
query= query&"preinscripciones.valor,preinscripciones.valor_total,preinscripciones.costo,preinscripciones.tipo_empresa, "
query= query&"preinscripciones.id_otic,preinscripciones.con_otic "
query= query&" FROM preinscripciones "
query= query&" INNER JOIN PROGRAMA ON PROGRAMA.ID_PROGRAMA=preinscripciones.id_programacion "
query= query&" INNER JOIN CURRICULO ON CURRICULO.ID_MUTUAL=PROGRAMA.ID_MUTUAL "
query= query&" INNER JOIN EMPRESAS ON EMPRESAS.ID_EMPRESA=preinscripciones.id_empresa "
query= query&" where preinscripciones.id_preinscripcion="&vid

set rsEmp = conn.execute (query)

if not rsEmp.eof and not rsEmp.bof then 
%>
 <form name="frmAutorizacion" id="frmAutorizacion" action="inspendientes/modificar.asp" method="post">
   <table cellspacing="0" cellpadding="1" border=0>
     <tr>
       <%
       if(rsEmp("tipo_empresa")="1")then
	   %>
       <td>Rut Empresa :</td>
       <%
	   end if
	   if(rsEmp("tipo_empresa")="2")then
	   %>
       <td>Rut Mandante :</td>
       <%
	   end if
	   %>
       <td><%=replace(FormatNumber(mid(rsEmp("RUT"), 1,len(rsEmp("RUT"))-2),0)&mid(rsEmp("RUT"), len(rsEmp("RUT"))-1,len(rsEmp("RUT"))),",",".")%></td>
       <td><input type="HIDDEN" id="Empresa" name="Empresa" value="<%=rsEmp("ID_EMPRESA")%>"/></td>
       <td>Raz&oacute;n Social : </td>
       <td colspan="4"><%=rsEmp("R_SOCIAL")%></td>
    </tr>
    <tr>
       <td>C&oacute;digo Curso :</td>
       <td><%=rsEmp("CODIGO")%></td>
       <td><input type="hidden" id="Programa" name="Programa" value="<%=rsEmp("ID_PROGRAMA")%>"/></td>
       <td>Nombre Curso :</td>
       <td colspan="4"><%=rsEmp("NOMBRE_CURSO")%></td>
    </tr>
    <tr>
       <td width="100">Fecha Inicio :</td>
       <td width="150"><%=rsEmp("FECHA_INICIO_")%></td>
       <td width="20">&nbsp;</td>
       <td width="105">Fecha Termino :</td>
 	   <td width="160"><%=rsEmp("FECHA_TERMINO")%></td>
       <td width="20">&nbsp;</td>
       <td width="145">&nbsp;</td>
       <td width="100">&nbsp;</td>
     </tr>
    <tr>
       <td colspan="2">Utiliza Franquicia Sence? :&nbsp;
       	   <input type="radio" name="COfran" id="COfran" value="1" onclick="MuestraFilas(1);"/>Si
           <input type="radio" name="COfran" id="COfran" value="0" onclick="MuestraFilas(0);"/>No</td>
       <td><input type="hidden" id="ConOtic" name="ConOtic" value="0"/></td>
       <td id="Oticpreg" style="display:none">Con OTIC :</td>       
       <% if(rs("ID_OTIC")<>"0")then %>     
       <td id="pregOtic" style="display:none">
       	  <input type="radio" name="COtic" id="COtic1" value="1" onclick="$('#ConOtic').val(this.value);$('#Mandante').show();"/>Si
          <input type="radio" name="COtic" id="COtic0" value="0" onclick="$('#ConOtic').val(this.value);$('#Mandante').hide();"/>No</td>
       <% else %>
       <td id="pregOtic" style="display:none">
       		<input type="radio" name="COtic" id="COtic1" value="1" disabled="disabled"/>Si
            <input type="radio" name="COtic" id="COtic0" value="0"/>No</td>
       <% end if %>
       <td>&nbsp;</td>
       <td id="Regpreg" style="display:none">N&deg; de Registro Sence:</td>
       <td id="RegFran" style="display:none"><input type="text" id="NReg_Fran" name="NReg_Fran" maxlength="50" size="20"/></td>
     </tr> 
	<%
		 if(rsEmp("id_otic")<>"0")then
			 set rsOtic = conn.execute ("select ID_EMPRESA,RUT,R_SOCIAL from EMPRESAS where EMPRESAS.ID_EMPRESA="&rsEmp("id_otic"))
			 %>
               <tr name="Mandante" id="Mandante" style.display ="">
                   <td>Rut OTIC :</td>
               <td><%=replace(FormatNumber(mid(rsOtic("RUT"), 1,len(rsOtic("RUT"))-2),0)&mid(rsOtic("RUT"), len(rsOtic("RUT"))-1,len(rsOtic("RUT"))),",",".")%></td>
                   <td></td>
                   <td>Raz&oacute;n Social :</td>
                   <td colspan="4"><label id="txtRsocOtic" name="txtRsocOtic"><%=rsOtic("R_SOCIAL")%></label></td>
               </tr>
    <%
	    end if
	%>
     <tr>
       <td>&nbsp;</td>
       <td></td>
       <td><input type="hidden" id="txtpreinscripcion" name="txtpreinscripcion" value="<%=rsEmp("id_preinscripcion")%>"/></td>
       <td><input type="hidden" id="txtIdOtic" name="txtIdOtic" value="<%=rsEmp("id_otic")%>"/>
       <input type="hidden" id="txtValorE" name="txtValorE" value="<%=rsEmp("costo")%>"/>
       <input type="hidden" id="txtOrdenCompraE" name="txtOrdenCompraE" value="<%=rsEmp("numero_compromiso")%>"/></td>
       <td></td>
    </tr>
   </table>
	<div width="200"><table id="list2"></table> 
   		<div id="pager2"></div> 
   </div>
    <table cellspacing="0" cellpadding="1" border=0>   
     <tr>
	   <td width="155">&nbsp;</td>
       <td width="145"></td>
       <td width="20"></td>
       <td width="155"></td>
       <td width="140"></td>
       <td width="20"></td>
       <td width="160"></td>
       <td width="145"></td>
      </tr>
       <tr>
       <td>N&#176; de Participantes : </td>
       <td><%=rsEmp("participantes")%></td>
	   <td>&nbsp;</td>
       <td>Valor por Alumno : </td>
       <td><label id="txtValor" name="txtValor"><%="$ "&replace(FormatNumber(rsEmp("VALOR"),0),",",".")%></label></td>
       <td>&nbsp;</td>
       <td>Valor Total :</td>
       <td><%="$ "&replace(FormatNumber(rsEmp("valor_total"),0),",",".")%></td>
     </tr>
      <tr>
       <td>Tipo Compromiso : </td>
       <%
	   tipo=""
	   
	   if(rsEmp("tipo_compromiso")=0)then
	   tipo="Orden de Compra"
	   end if
	   
	   if(rsEmp("tipo_compromiso")=1)then
	   tipo="Carta Inscripción"
	   end if
	   
	   if(rsEmp("tipo_compromiso")=2)then
	   	tipo="Depósito Cheque"
	   end if
	   
	    if(rsEmp("tipo_compromiso")=3)then
	   	tipo="Transferencia"
	   end if
	   
	   if(rsEmp("tipo_compromiso")=4)then
	   	tipo="Carta Compromiso"
	   end if
	   
	   %>
       
     <td><%=tipo%></td>
	   <td>&nbsp;</td>
       <td>N&uacute;mero OC/CI/DEP : </td>
       <td><%=rsEmp("numero_compromiso")%></td>
       <td>&nbsp;</td>
         <% if(rsEmp("ID_OTIC")<>"0")then %>     
       <td>Curso Inscrito por OTIC :</td>
       <td>
       <%
	   checkedSi=""
	   checkedNo=""
	   estadoCOtic=0
	   
	   if(rsEmp("con_otic")="1")then
		   checkedSi="checked='checked'"
		   estadoCOtic=1
	   else
		   checkedNo="checked='checked'"
		   estadoCOtic=0
	   end if
	   %>
           <input type="radio" name="COtic" id="COtic" value="1" <%=checkedSi%> onclick="$('#ConOtic').val(this.value);"/>Si
           <input type="radio" name="COtic" id="COtic" value="0" <%=checkedNo%> onclick="$('#ConOtic').val(this.value);"/>No
           <input type="hidden" id="ConOtic" name="ConOtic" value="<%=estadoCOtic%>"/></td>
       <% else %>
        <td colspan="2">&nbsp;
        <input type="hidden" id="ConOtic" name="ConOtic" value="0"/></td>
       <% end if %>
     </tr>
       <tr>
       <td> Documento  : </td>
       <td colspan="7"><a href="#" onclick="documento('<%=rsEmp("doc")%>');">Ver Documento</a></td>
     </tr>
     <tr>
      <td colspan="8">&nbsp;<input type="hidden" id="documentos" name="documentos" value="<%=rsEmp("doc")%>"/></td>
     </tr>
     <tr>
    	<td colspan="8"><label>
    	<input type="checkbox" name="terminos" id="terminos" tabindex="17"/>
    	Rechazar Inscripci&oacute;n</label></td>
   	  </tr>
      <tr>
        <td colspan="8">&nbsp;</td>
      </tr>
   </table>
</form> 
<%
end if
%>