<%
on error resume next
%>
<!--#include file="../conexion.asp"-->
<%
Response.CodePage = 65001
Response.CharSet = "utf-8"
vid = Request("id")

dim query
query= "SELECT dbo.MayMinTexto(CURRICULO.NOMBRE_CURSO) as NOMBRE_CURSO,CONVERT(VARCHAR(10),PROGRAMA.FECHA_INICIO_, 105) as FECHA_INICIO_, "
query= query&"CONVERT(VARCHAR(10),PROGRAMA.FECHA_TERMINO, 105) as FECHA_TERMINO, "
query= query&"EMPRESAS.R_SOCIAL,EMPRESAS.RUT,preinscripciones.id_otic, PROGRAMA.ID_PROGRAMA,EMPRESAS.ID_EMPRESA,"
query= query&"CURRICULO.CODIGO,CURRICULO.SENCE,preinscripciones.tipo_compromiso,preinscripciones.numero_compromiso, "
query= query&"preinscripciones.id_preinscripcion,preinscripciones.doc_compromiso as doc,preinscripciones.participantes,"
query= query&"preinscripciones.valor,preinscripciones.valor_total,preinscripciones.costo,preinscripciones.tipo_empresa, "
query= query&"preinscripciones.id_otic,preinscripciones.con_otic,preinscripciones.con_franquicia,preinscripciones.n_reg_sence, "
query= query&"preinscripciones.id_proyecto,preinscripciones.id_tipo_venta,PROGRAMA.ID_MUTUAL,"
query= query&"b.nombre_banco,cc.numero_cuenta,preinscripciones.monto_transferencia,fecha_trans=CONVERT(VARCHAR(10),preinscripciones.fecha_transferencia, 105), "
query= query&"preinscripciones.rut_depositante,preinscripciones.nombre_depositante,preinscripciones.email_depositante,preinscripciones.relacion_depositante FROM preinscripciones "
query= query&" INNER JOIN PROGRAMA ON PROGRAMA.ID_PROGRAMA=preinscripciones.id_programacion "
query= query&" INNER JOIN CURRICULO ON CURRICULO.ID_MUTUAL=PROGRAMA.ID_MUTUAL "
query= query&" INNER JOIN EMPRESAS ON EMPRESAS.ID_EMPRESA=preinscripciones.id_empresa "
query= query&" left join CUENTAS_CORRIENTES cc on cc.id_cuenta_corriente=preinscripciones.id_cuenta_corriente and cc.id_banco=preinscripciones.id_banco "
query= query&" left join BANCOS b on b.id_banco=preinscripciones.id_banco "
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
       <td><input type="hidden" id="Programa" name="Programa" value="<%=rsEmp("ID_PROGRAMA")%>"/><input type="hidden" id="id_proyecto" name="id_proyecto" value="<%=rsEmp("id_proyecto")%>"/></td>
       <td>Nombre Curso :</td>
       <td colspan="4"><%=rsEmp("NOMBRE_CURSO")%><%if(rsEmp("ID_MUTUAL")="39" or rsEmp("ID_MUTUAL")="40")then%>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font color="#FF0000">Pertenece a Proyecto:</font>&nbsp;<%=rsEmp("id_proyecto")%><%end if%></td>
    </tr>
     <tr>
       <td width="100">Fecha Inicio :</td>
       <td width="150"><%=rsEmp("FECHA_INICIO_")%></td>
       <td width="20">&nbsp;</td>
       <td width="105">Fecha Termino :</td>
 	   <td width="160"><%=rsEmp("FECHA_TERMINO")%></td>
       <td width="20">&nbsp;</td>
       <td width="160">&nbsp;</td>
       <td width="180">&nbsp;</td>
     </tr>
     <tr>
       <td colspan="2">Utiliza Franquicia Sence? :&nbsp;
       <%
	   ckdFranSi=""
	   ckdFranNo=""
   
	   if(rsEmp("con_franquicia")="1")then
		  ckdFranSi="checked='checked'"
	   else
		  ckdFranNo="checked='checked'"
	   end if
	   %>          
       	   <input type="radio" name="COfran" id="COfran" value="1" <%=ckdFranSi%> disabled="disabled"/>Si
           <input type="radio" name="COfran" id="COfran" value="0" <%=ckdFranNo%> disabled="disabled"/>No</td>
       <td><input type="hidden" id="Confran" name="Confran" value="<%=rsEmp("con_franquicia")%>"/></td>
       <td id="Oticpreg" style="display:none">Con OTIC :</td>   
       <%
	   checkedSi=""
	   checkedNo=""
   
	   if(rsEmp("con_otic")="1")then
		   checkedSi="checked='checked'"
	   else
		   checkedNo="checked='checked'"
	   end if
	   %>           
       <td id="pregOtic" style="display:none">
       	  <input type="radio" name="COtic" id="COtic1" value="1" <%=checkedSi%> disabled="disabled"/>Si
          <input type="radio" name="COtic" id="COtic0" value="0" <%=checkedNo%> disabled="disabled"/>No
          <input type="hidden" id="ConOtic" name="ConOtic" value="<%=rsEmp("con_otic")%>"/></td>
       <td><input type="hidden" id="Reg_Fran" name="Reg_Fran" value="<%=rsEmp("n_reg_sence")%>"/></td>
       <td id="Regpreg" style="display:none">N&deg; de Registro Sence:</td>
       <td id="RegFran" style="display:none"><%=rsEmp("n_reg_sence")%></td>
     </tr> 
         <tr>
           <td>Tipo Venta:<input type="hidden" id="tipoVenta" name="tipoVenta" value="<%=rsEmp("id_tipo_venta")%>"/></td>
           <% if(rsEmp("id_tipo_venta")="1")then %>   
        <td>Venta Directa</td>
        <% elseif(rsEmp("id_tipo_venta")="2") then%>
        <td>Venta Broker</td>
     <% else %>
     <td>Sin Asignaci&oacute;n</td>
     <% end if %>
         </tr>     
	<%
		 if(rsEmp("con_franquicia")="1" and rsEmp("con_otic")="1")then
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
	   tipo="Vale Vista"
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
	   

	   if(rsEmp("tipo_compromiso")=5)then
	   	tipo="EDP"
	   end if
	   %>
       
       <td><%=tipo%></td>
	   <td>&nbsp;</td>
       <td>N&uacute;mero OC/CI/DEP : </td>
       <td><%=rsEmp("numero_compromiso")%></td>
       <td>&nbsp;</td>
       <td>Documento  : </td>
       <td><a href="#" onclick="documento('<%=rsEmp("doc")%>');">Ver Documento</a></td>
     </tr>
<%
if(rsEmp("tipo_compromiso")=2 or rsEmp("tipo_compromiso")=3)then
%>
    <tr>
      <td colspan="2">Banco : <%=rsEmp("nombre_banco")%></td>
      <td>&nbsp;</td>
      <td>N&uacute;mero Cuenta : </td>
      <td><%=rsEmp("numero_cuenta")%></td>
      <td></td>
      <td>Rut Depositante : </td>
      <td><%=rsEmp("rut_depositante")%></td>
    </tr>
    <tr>
      <td>Monto: </td>
      <td><%=rsEmp("monto_transferencia")%></td>
      <td>&nbsp;</td>
      <td>Fecha D&eacute;posito: </td>
      <td><%=rsEmp("fecha_trans")%></td>
      <td></td>
      <td>Nombre Depositante : </td>
      <td><%=rsEmp("nombre_depositante")%></td>
    </tr>
    <tr>
      <td>Relaci&oacute;n Depositante con Empresa : </td>
      <td><%=rsEmp("relacion_depositante")%></td>
      <td></td>
      <td>Mail de Depositante : </td>
      <td><%=rsEmp("email_depositante")%></td>
      <td>&nbsp;</td>
      <td></td>
      <td></td>
    </tr>
<%
end if
%>
     <tr>
      <td colspan="8">&nbsp;<input type="hidden" id="documentos" name="documentos" value="<%=rsEmp("doc")%>"/></td>
     </tr>
     <tr>
    	<td colspan="8"><!--<label>
    	<input type="checkbox" name="terminos" id="terminos" tabindex="17"/>
    	Rechazar Inscripci&oacute;n</label>--></td>
   	  </tr>
      <tr>
        <td colspan="8">&nbsp;</td>
      </tr>
   </table>
</form> 
<%
end if
%>