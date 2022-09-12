<%
on error resume next
%>
<!--#include file="../conexion.asp"-->
<%
Response.CodePage = 65001
Response.CharSet = "utf-8"
vid = Request("id")
dim query
query= "select [RUT_EMPRESA],[NOMBRE_EMPRESA],[NOMBRE_ARCHIVO],[CreditoPlazo],[CreditoSol],[CreditoDia],[CreditoEncargado1],"&_
       "[CreditoCorreo1],[CreditoFono1],[CreditoEncargado2],[CreditoCorreo2],[CreditoFono2],[CreditoRepLegal],[CreditoRepLegalCorreo],"&_
       "[CreditoRepLegalFono],[CreditoHorario],[CreditoInfo],[cond_OC],[cond_HES],[cond_MIGO],[cond_OTRO],[cond_OTROTxt]"&_
       " from SOLICITUDES_CREDITO where ID_SOLICITUD="&vid

set rsEmp = conn.execute (query)

set rsDetEmp = conn.execute ("select ID_EMPRESA,RUT,R_SOCIAL from EMPRESAS where ID_EMPRESA='"&Request("idEmpresa")&"'")


if not rsEmp.eof and not rsEmp.bof then 
%>
			<form name="frmSolicitud" id="frmSolicitud" action="solicitud/solicitudcreditoAjax.asp" method="post">
<table cellspacing="3" cellpadding="1" border=0>
				<tr>
					<td width="300"></td>
					<td width="20"></td>
					<td width="300"></td>
					<td width="20"></td>
					<td width="300"></td> 
				</tr>
				<tr>
					<td colspan="2"><b>Rut Empresa:</b> <%=rsEmp("RUT_EMPRESA")%></td>
					<td colspan="3"><b>Nombre Empresa:</b>&nbsp;<%=rsEmp("NOMBRE_EMPRESA")%></td>
				</tr>
				<tr>
					<td colspan="5">Linea de Credito Solicitado: <%=rsEmp("CreditoSol")%></td>
				</tr>
				<tr>
					<td colspan="2">Plazo De Pago Solicitado:&nbsp;&nbsp;&nbsp;&nbsp;<%=rsEmp("CreditoPlazo")%></td>
					<td colspan="3">Día De Pago : <%=rsEmp("CreditoDia")%></td>
				</tr>
	 			<tr>
					<td colspan="5">&nbsp;</td>
				</tr>
				<tr>
					<td colspan="5"><center>
					<h3><em style="text-transform: capitalize;">Información Solicitada</em></h3></center></td>
				</tr>
				<tr>
					<td colspan="5">&nbsp;</td>
				</tr>
				<tr>
					<td>Encargado Compras : <%=rsEmp("CreditoEncargado1")%></td>
					<td></td>
					<td>Correo Electrónico : <%=rsEmp("CreditoCorreo1")%></td>
					<td></td>
					<td>Teléfono : <%=rsEmp("CreditoFono1")%></td>
				</tr>
				<tr>
					<td>Encargado Pagos : <%=rsEmp("CreditoEncargado2")%></td>
					<td></td>
					<td>Correo Electrónico : <%=rsEmp("CreditoCorreo2")%></td>
					<td></td>
					<td>Teléfono : <%=rsEmp("CreditoFono2")%></td>
				</tr>
				<tr>
					<td colspan="2">Observaciones (Horario De Atención A Proveedores) :</td>
					<td colspan="3"><%=rsEmp("CreditoHorario")%></td>
				</tr>
				<tr>
					<td>Representante Legal : <%=rsEmp("CreditoRepLegal")%></td>
					<td></td>
					<td>Correo Electrónico : <%=rsEmp("CreditoRepLegalCorreo")%></td>
					<td></td>
					<td>Teléfono : <%=rsEmp("CreditoRepLegalFono")%></td>
				</tr>
<%
checkoc=""
if(rsEmp("cond_OC")="1")then
	checkoc="checked"
end if

checkhes=""
if(rsEmp("cond_HES")="1")then
	checkhes="checked"
end if

checkmigo=""
if(rsEmp("cond_MIGO")="1")then
	checkmigo="checked"
end if

checkotro=""
if(rsEmp("cond_OTRO")="1")then
	checkotro="checked"
end if

%>
				<tr>
					<td colspan="2">Condiciones Particulares Para Facturar, Señale Requisitos Especiales:</td>
					<td colspan="3"><label>Se Requiere Oc:</label><input type="checkbox" name="cond_OC" id="cond_OC" value="1" disabled="disabled" <%=checkoc%>>
							<label>Se Requiere Hes:</label><input type="checkbox" name="cond_HES" id="cond_HES" value="1" disabled="disabled" <%=checkhes%>></br>
							<label>Se Requiere Migo:</label><input type="checkbox" name="cond_MIGO" id="cond_MIGO" value="1" disabled="disabled" <%=checkmigo%>>
							<label>Indique Otro:</label><input type="checkbox" name="cond_OTRO" id="cond_OTRO" value="1" disabled="disabled" <%=checkotro%>>
										    
<%
if(rsEmp("cond_OTRO")="1")then
%>
Detalle Otro: <%=rsEmp("cond_OTROTxt")%>
<%
end if
%>
</td>
				</tr>
				<tr>
					<td colspan="2">Información Adicional (Señale Observaciones Para Considerar) :</td>
					<td colspan="3"><%=rsEmp("CreditoInfo")%></td>
				</tr>
				<tr>
					<td colspan="5"><center>
					<h3><em style="text-transform: capitalize;">Documentos</em></h3></center></td>
				</tr>
				<tr>
					<td colspan="5">&nbsp;</td>
				</tr>
				<tr>
					<td colspan="5">Documento : <a href="#" onclick="documento('<%=rsEmp("NOMBRE_ARCHIVO")%>');">Ver Documento</a></td>
				</tr>
     				
			</table>
			</form> 
<%
   else
%> 
			<form name="frmSolicitud" id="frmSolicitud" action="solicitud/solicitudcreditoAjax.asp" method="post">
				<table cellspacing="3" cellpadding="1" border=0>
				<tr>
					<td width="370"></td>
					<td width="20"></td>
					<td width="300"></td>
					<td width="20"></td>
					<td width="300"><input type="hidden" id="solId" name="solId" /> <input type="hidden" id="operacion" name="operacion" /></td> 
				</tr>
				<tr>
					<td colspan="2"><b>Rut Empresa:</b> <%=rsDetEmp("RUT")%><input id="txRut" name="txRut" type="hidden" value="<%=rsDetEmp("RUT")%>" /></td>
					<td colspan="3"><b>Nombre Empresa:</b> <%=rsDetEmp("R_SOCIAL")%><input id="txtNombreEmpresa" name="txtNombreEmpresa" type="hidden" value="<%=rsDetEmp("R_SOCIAL")%>"/></td>


				</tr>
				<tr>
					<td colspan="5">Linea de Credito Solicitado: <input id="txtCreditoSol" name="txtCreditoSol" type="text" size="15"/></td>
				</tr>
				<tr>
					<td colspan="2">Plazo De Pago Solicitado:&nbsp;&nbsp;&nbsp;&nbsp;<select id="txtCreditoPlazo" name="txtCreditoPlazo" tabindex="9">
							<option value="">Seleccione</option>
							<option value="Contado">Contado</option>
							<option value="30 Días">30 Dias</option>
							</select></td>
					<td colspan="3">Día De Pago : <input id="txtCreditoDia" name="txtCreditoDia" type="text" size="15"/></td>
				</tr>
	 			<tr>
					<td colspan="5">&nbsp;</td>
				</tr>
				<tr>
					<td colspan="5"><center>
					<h3><em style="text-transform: capitalize;">Información Solicitada</em></h3></center></td>
				</tr>
				<tr>
					<td colspan="5">&nbsp;</td>
				</tr>
				<tr>
					<td>Encargado Compras : <input id="txtCreditoEncargado1" name="txtCreditoEncargado1" type="text"/></td>
					<td></td>
					<td>Correo Electrónico : <input id="txtCreditoCorreo1" name="txtCreditoCorreo1" type="text" size="15"/></td>
					<td></td>
					<td>Teléfono : <input id="txtCreditoFono1" name="txtCreditoFono1" type="text" size="15"/></td>
				</tr>
				<tr>
					<td>Encargado Pagos : <input id="txtCreditoEncargado2" name="txtCreditoEncargado2" type="text" size="15"/></td>
					<td></td>
					<td>Correo Electrónico : <input id="txtCreditoCorreo2" name="txtCreditoCorreo2" type="text" size="15"/></td>
					<td></td>
					<td>Teléfono : <input id="txtCreditoFono2" name="txtCreditoFono2" type="text" size="15"/></td>
				</tr>
				<tr>
					<td colspan="2">Observaciones (Horario De Atención A Proveedores) :</td>
					<td colspan="3"><textarea type="text" rows="3" cols="50" id="txtCreditoHorario" tabindex="3" name="txtCreditoHorario"></textarea></td>
				</tr>
				<tr>
					<td>Representante Legal : <input id="txtCreditoRepLegal" name="txtCreditoRepLegal" type="text" size="15"/></td>
					<td></td>
					<td>Correo Electrónico : <input id="txtCreditoRepLegalCorreo" name="txtCreditoRepLegalCorreo" type="text" size="15"/></td>
					<td></td>
					<td>Teléfono : <input id="txtCreditoRepLegalFono" name="txtCreditoRepLegalFono" type="text" size="15"/></td>
				</tr>
				<tr>
					<td colspan="2">Condiciones Particulares Para Facturar, Señale Requisitos Especiales:</td>
					<td colspan="3"><label>Se Requiere Oc:</label><input type="checkbox" name="cond_OC" id="cond_OC" value="1" onclick="asignaChk('cond_OC');">
							<label>Se Requiere Hes:</label><input type="checkbox" name="cond_HES" id="cond_HES" value="1" onclick="asignaChk('cond_HES');"></br>
							<label>Se Requiere Migo:</label><input type="checkbox" name="cond_MIGO" id="cond_MIGO" value="1" onclick="asignaChk('cond_MIGO');">
							<label>Indique Otro:</label><input type="checkbox" name="cond_OTRO" id="cond_OTRO" value="1" onclick="asignaChk('cond_OTRO');">
										    <input id="cond_OTROTxt" name="cond_OTROTxt" type="text" size="15"/></td>
				</tr>
				<tr>
					<td colspan="2">Información Adicional (Señale Observaciones Para Considerar) :</td>
					<td colspan="3"><textarea type="text" rows="3" cols="50" id="txtCreditoInfo" tabindex="3" name="txtCreditoInfo"></textarea></td>
				</tr>
	                        <tr>
     					<td colspan="5"><center>&nbsp;<img id="loading" src="../images/loader.gif" style="display:none;"/></center></td>
     				</tr>
				<tr>
					<td colspan="5"><center>
					<h3><em style="text-transform: capitalize;">Documentos</em></h3></center></td>
				</tr>
				<tr>
					<td colspan="5">&nbsp;<input id="checkcond_OC" name="checkcond_OC" type="hidden"/>
							      <input id="checkcond_HES" name="checkcond_HES" type="hidden"/>
							      <input id="checkcond_MIGO" name="checkcond_MIGO" type="hidden"/>
							      <input id="checkcond_OTRO" name="checkcond_OTRO" type="hidden"/></td>
				</tr>
				<tr>
					<td colspan="1">Subir Documento : </td>
					<td colspan="4"><input type="file" name="txtDoc" id="txtDoc" tabindex="5" maxlength="400" size="25"></td>
				</tr>
     				
			</table>
			</form> 
<%
   end if
%>
<div id="messageBox1" style="height:100px;overflow:auto;width:600px;"> 
  	<ul></ul> 
</div> 
