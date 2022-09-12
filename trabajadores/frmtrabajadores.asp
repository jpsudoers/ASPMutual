<%
on error resume next
%>
<!--#include file="../conexion.asp"-->
<%
Response.CodePage = 65001
Response.CharSet = "utf-8"
vid = Request("id")
dim query
query= "select * from TRABAJADOR where ID_TRABAJADOR="&vid
set rsEmp = conn.execute (query)

if not rsEmp.eof and not rsEmp.bof then 
%>
 <form name="frmTrabajador" id="frmTrabajador" action="trabajadores/modificar.asp" method="post">
<table cellspacing="0" cellpadding="1" border=0>
	<tr>
    	<td width="110">Rut :</td>
      	<td width="200"><input id="txRut" name="txRut" type="text" tabindex="1" onKeyPress="" size="12" maxlength="11" value="<%=rsEmp("rut")%>" onfocus="estudios('<%=rsEmp("ESCOLARIDAD")%>')"/></td>
        <td width="20" ><input type="hidden" id="txtId" name="txtId" value="<%=rsEmp("id_trabajador")%>" /></td>
        <td width="150">&nbsp;</td>
        <td width="172">&nbsp;</td>
        <td width="20">&nbsp;</td>
        <td width="160">&nbsp;</td>
        <td width="131" >&nbsp;</td>
   	</tr>
     <tr>
        <td>Nombres :</td>
        <td><input id="txtNomb" name="txtNomb" type="text" tabindex="2" maxlength="50" value="<%=rsEmp("nombres")%>"  size="20" onkeypress="return val(event)"/></td>
        <td></td>
        <td>Apellido Paterno :</td>
        <td><input id="txtAp_pater" name="txtAp_pater" type="text" tabindex="3" maxlength="50" value="<%=rsEmp("a_paterno")%>"  size="20" onkeypress="return val(event)"/></td>
        <td></td>
        <td>Apellido Materno :</td>
        <td><input id="txtAp_mater" name="txtAp_mater" type="text" tabindex="4" maxlength="50" value="<%=rsEmp("a_materno")%>"  size="20" onkeypress="return val(event)"/></td>
    </tr>
     <tr>
    	<td>Fecha Nac :</td>
        <td><input id="txtFecha" name="txtFecha" type="text" tabindex="5" maxlength="50" size="12" value="<%=rsEmp("FECHA_NACIMIENTO")%>"/></td>
        <td></td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
    </tr>
      <tr>
        <td>Direcci&oacute;n :</td>
        <td><input id="txtDir" name="txtDir" type="text" tabindex="6" maxlength="20" value="<%=rsEmp("DIRECCION")%>" size="20"/></td>
        <td></td>
        <td>Comuna :</td>
        <td><input id="txtCom" name="txtCom" type="text" tabindex="7" size="20" maxlength="40" value="<%=rsEmp("comuna")%>" onkeypress="return val(event)"/></td>
        <td></td>
        <td>Ciudad :</td>
        <td><input id="txtCiu" name="txtCiu" type="text" tabindex="8"  size="20"  maxlength="30" value="<%=rsEmp("ciudad")%>" onkeypress="return val(event)"/></td>
    </tr>
    <tr>
        <td>Escolaridad :</td>
        <td><select id="txtEscol" name="txtEscol" tabindex="9">
        <option value="0">Sin Escolaridad</option>
        <option value="1">B&aacute;sica Incompleta</option>
        <option value="2">B&aacute;sica Completa</option>
        <option value="3">Media Incompleta</option>
        <option value="4">Media Completa</option>
        <option value="5">Superior T&eacute;cnica Incompleta</option>
        <option value="6">Superior T&eacute;cnica Profesional Completa </option>
        <option value="7">Universitaria Incompleta</option>
        <option value="8">Universitaria Completa</option>
        </select></td>
        <td></td>
        <td>Especialidad :</td>
        <td><input id="txtEsp" name="txtEsp" type="text" tabindex="10" size="20" maxlength="40" value="<%=rsEmp("especialidad")%>" onkeypress="return val(event)"/></td>
        <td></td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
    </tr>
      <tr>
    	<td>Telefono Fijo :</td>
        <td><input id="txtFon" name="txtFon" type="text" tabindex="11" maxlength="50" value="<%=rsEmp("fono_fijo")%>"  size="12" onKeyPress="return acceptNum(event)"/></td>
        <td></td>
        <td>Celular :</td>
        <td><input id="txtCel" name="txtCel" type="text" tabindex="12" maxlength="50" value="<%=rsEmp("celular")%>"  size="12" onKeyPress="return acceptNum(event)"/></td>
        <td></td>
        <td>Email :</td>
        <td><input id="txtMail" name="txtMail" type="text" tabindex="13" maxlength="50" value="<%=rsEmp("email")%>"  size="20"/></td>
    </tr>
	<tr>
    	<td>Cargo :</td>
        <td><input id="txtCargo" name="txtCargo" type="text" tabindex="14" maxlength="50" value="<%=rsEmp("CARGO_EMPRESA")%>"  size="12"  onkeypress="return val(event)"/></td>
        <td></td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
    </tr>
</table>
</form> 
<%
   else
%> <form name="frmTrabajador" id="frmTrabajador" action="trabajadores/insertar.asp" method="post">
<table cellspacing="0" cellpadding="1" border=0>
	<tr>
    	<td width="102">Rut :</td>
      	<td width="166"><input id="txRut" name="txRut" type="text" tabindex="1" onKeyPress="" size="12" maxlength="11" /></td>
        <td width="20" ></td>
        <td width="126">&nbsp;</td>
        <td width="172">&nbsp;</td>
        <td width="20">&nbsp;</td>
        <td width="128">&nbsp;</td>
        <td width="131" >&nbsp;</td>
   	</tr>
     <tr>
        <td>Nombres :</td>
        <td><input id="txtNomb" name="txtNomb" type="text" tabindex="2" maxlength="50" size="20" onkeypress="return val(event)"/></td>
        <td></td>
        <td>Apellido Paterno :</td>
        <td><input id="txtAp_pater" name="txtAp_pater" type="text" tabindex="3" maxlength="50" size="20" onkeypress="return val(event)"/></td>
        <td></td>
        <td>Apellido Materno :</td>
        <td><input id="txtAp_mater" name="txtAp_mater" type="text" tabindex="4" maxlength="50" size="20" onkeypress="return val(event)"/></td>
    </tr>
     <tr>
    	<td>Fecha Nac :</td>
        <td><input id="txtFecha" name="txtFecha" type="text" tabindex="5" maxlength="50" size="12"/></td>
        <td></td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
    </tr>
      <tr>
        <td>Direcci&oacute;n :</td>
        <td><input id="txtDir" name="txtDir" type="text" tabindex="6" maxlength="20" value="" size="20"/></td>
        <td></td>
        <td>Comuna :</td>
        <td><input id="txtCom" name="txtCom" type="text" tabindex="7" size="20" maxlength="40" value="" onkeypress="return val(event)" /></td>
        <td></td>
        <td>Ciudad :</td>
        <td><input id="txtCiu" name="txtCiu" type="text" tabindex="8"  size="20"  maxlength="30" value="" onkeypress="return val(event)"/></td>
    </tr>
    <tr>
        <td>Escolaridad :</td>
        <td><select id="txtEscol" name="txtEscol" tabindex="9">
        <option value="0" selected="selected">Sin Escolaridad</option>
        <option value="1">B&aacute;sica Incompleta</option>
        <option value="2">B&aacute;sica Completa</option>
        <option value="3">Media Incompleta</option>
        <option value="4">Media Completa</option>
        <option value="5">Superior T&eacute;cnica Incompleta</option>
        <option value="6">Superior T&eacute;cnica Profesional Completa </option>
        <option value="7">Universitaria Incompleta</option>
        <option value="8">Universitaria Completa</option>        </select></td>
        <td></td>
        <td>Especialidad :</td>
        <td><input id="txtEsp" name="txtEsp" type="text" tabindex="10" size="20" maxlength="40" value="" onkeypress="return val(event)"/></td>
        <td></td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
    </tr>
      <tr>
    	<td>Telefono Fijo :</td>
        <td><input id="txtFon" name="txtFon" type="text" tabindex="11" maxlength="50" value="" size="12" /></td>
        <td></td>
        <td>Celular :</td>
        <td><input id="txtCel" name="txtCel" type="text" tabindex="12" maxlength="50" value="" size="12" /></td>
        <td></td>
        <td>Email :</td>
        <td><input id="txtMail" name="txtMail" type="text" tabindex="13" maxlength="50" value=""  size="20"/></td>
    </tr>
	<tr>
    	<td>Cargo :</td>
        <td><input id="txtCargo" name="txtCargo" type="text" tabindex="14" maxlength="50" size="12" onkeypress="return val(event)"/></td>
        <td></td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
    </tr>
</table>
</form> 
<%
   end if
%>
<div id="messageBox1" style="height:100px;overflow:auto;width:300px;"> 
  	<ul></ul> 
</div> 
