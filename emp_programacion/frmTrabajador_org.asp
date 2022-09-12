<%
on error resume next
%>
<!--#include file="../conexion.asp"-->
<%
Response.CodePage = 65001
Response.CharSet = "utf-8"

query="select TRABAJADOR.ID_TRABAJADOR,TRABAJADOR.RUT,TRABAJADOR.NOMBRES,TRABAJADOR.CARGO_EMPRESA,TRABAJADOR.ESCOLARIDAD,"
query= query&"TRABAJADOR.NOM_TRAB,TRABAJADOR.APATERTRAB,TRABAJADOR.AMATERTRAB,NACIONALIDAD,ID_EXTRANJERO "
query= query&" from PREINSCRIPCION_TRABAJADOR "
query= query&" inner join TRABAJADOR on TRABAJADOR.ID_TRABAJADOR=PREINSCRIPCION_TRABAJADOR.id_trabajador "
query= query&" where PREINSCRIPCION_TRABAJADOR.preinscripcionTemp='"&Request("preinsId")&"'"
query= query&" and PREINSCRIPCION_TRABAJADOR.id_trabajador='"&Request("trabId")&"'"

set rsEmp = conn.execute (query)

if not rsEmp.eof and not rsEmp.bof then 
%>
 <form name="frmTrabajador" id="frmTrabajador" action="emp_programacion/modificar_tab.asp" method="post">
 <table cellspacing="0" cellpadding="1" border=0>
 <% 
 dim desNacionalidad
 desNacionalidad="Chileno(a)"
 
 if(rsEmp("NACIONALIDAD")="1")then
 	desNacionalidad="Extranjero(a)"
 end if
 %>
     <tr>
       <td>Nacionalidad : </td>
       <td><%=desNacionalidad%></td>
     </tr>
     <%if(rsEmp("NACIONALIDAD")="0")then%>
      <tr>
       <td>Rut : </td>
       <td><input id="txtRutTrab" name="txtRutTrab" type="hidden" tabindex="20" maxlength="12" size="12" disabled="disabled" value="<%=rsEmp("RUT")%>"/><%=rsEmp("RUT")%></td>
     </tr>
     <%else%>
     <tr>
       <td>Doc. Identificaci&oacute;n :</td>
       <td><input id="txtPasTrab" name="txtPasTrab" type="hidden" tabindex="22" maxlength="12" size="12" disabled="disabled" value="<%=rsEmp("ID_EXTRANJERO")%>"/><%=rsEmp("ID_EXTRANJERO")%></td>
     </tr>
     <%end if%>
      <tr>
       <td width="130">Nombres  : </td>
       <td width="250"><input id="txtNomTrab" name="txtNomTrab" type="text" tabindex="21" maxlength="30" size="30" value="<%=rsEmp("NOM_TRAB")%>"/></td>
     </tr>
     <tr>
       <td>Apellido Paterno  : </td>
       <td><input id="txtAPaterTrab" name="txtAPaterTrab" type="text" tabindex="23" maxlength="30" size="30" value="<%=rsEmp("APATERTRAB")%>"/></td>
     </tr>
     <tr>
       <td>Apellido Materno  : </td>
       <td><input id="txtAMaterTrab" name="txtAMaterTrab" type="text" tabindex="24" maxlength="30" size="30" value="<%=rsEmp("AMATERTRAB")%>"/></td>
     </tr>
     <tr>
       <td>Cargo  : </td>
       <td><input id="txtCargoTrab" name="txtCargoTrab" type="text" tabindex="25" maxlength="50" size="30" value="<%=rsEmp("CARGO_EMPRESA")%>"/></td>
     </tr>
     <tr>
       <td>Escolaridad : </td>
       <td><select id="escolaridadTrab" name="escolaridadTrab" tabindex="26"></select>
       <input id="txtIdEscolaridad" name="txtIdEscolaridad" type="hidden" value="<%=rsEmp("ESCOLARIDAD")%>"/>
       <input id="txtTrabID" name="txtTrabID" type="hidden" value="<%=rsEmp("ID_TRABAJADOR")%>"/></td>
     </tr>
     <tr>
       <td>&nbsp;</td>
       <td>&nbsp;</td>
     </tr>
   </table>
</form> 
<%
   else
%><form name="frmTrabajador" id="frmTrabajador" action="emp_programacion/insertar_tab.asp" method="post">
 <table cellspacing="0" cellpadding="1" border=0>
     <tr>
       <td width="130">Nacionalidad : </td>
       <td width="250"><input type="checkbox" name="extranjero" id="extranjero" tabindex="21" onclick="cambiarCampo();"/>Extranjero
       </td>
     </tr>
     <tr name="rutTrab" id="rutTrab">
       <td>Rut : </td>
       <td><input id="txtRutTrab" name="txtRutTrab" type="text" tabindex="22" maxlength="12" size="12" onblur="datosTrabajador(this.value);"/> 
     Ejemplo : 16213589-k</td>
     </tr>
     <tr name="pasTrab" id="pasTrab" style="display:none">
       <td>Doc. Identificaci&oacute;n :</td>
       <td><input id="txtPasTrab" name="txtPasTrab" type="text" tabindex="22" maxlength="30" size="30" value="1"/></td>
     </tr>
      <tr>
       <td>Nombres  : </td>
       <td><input id="txtNomTrab" name="txtNomTrab" type="text" tabindex="23" maxlength="30" size="30"/></td>
     </tr>
     <tr>
       <td>Apellido Paterno  : </td>
       <td><input id="txtAPaterTrab" name="txtAPaterTrab" type="text" tabindex="24" maxlength="30" size="30"/></td>
     </tr>
     <tr>
       <td>Apellido Materno  : </td>
       <td><input id="txtAMaterTrab" name="txtAMaterTrab" type="text" tabindex="25" maxlength="30" size="30"/></td>
     </tr>
     <tr>
       <td>Cargo  : </td>
       <td><input id="txtCargoTrab" name="txtCargoTrab" type="text" tabindex="26" maxlength="50" size="30"/></td>
     </tr>
     <tr>
       <td>Escolaridad : </td>
       <td><select id="escolaridadTrab" name="escolaridadTrab" tabindex="27"></select>
       <input id="txtIdEscolaridad" name="txtIdEscolaridad" type="hidden"/>
       <input id="txtTrabID" name="txtTrabID" type="hidden" value="0"/>
       <input id="txtTrabPreins" name="txtTrabPreins" type="hidden" value="<%=Request("preinsId")%>"/>
       <input type="hidden" id="insGrilla" name="insGrilla" value="1"/>
       <input type="hidden" id="checkExtran" name="checkExtran" value="0"/>
       </td>
     </tr>
     <tr>
       <td>&nbsp;</td>
       <td>&nbsp;</td>
     </tr>
   </table>
</form> 
<%
   end if
%>
<div id="messageBox2" style="height:90px;overflow:auto;width:350px;"> 
  	<ul></ul> 
</div> 