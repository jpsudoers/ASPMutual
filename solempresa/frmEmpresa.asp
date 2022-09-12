<!--#include file="../conexion.asp"-->
<%
Response.CodePage = 65001
Response.CharSet = "utf-8"
vid = Request("id")
dim query
query= "select * from EMPRESAS where id_empresa="&vid
set rsEmp = conn.execute (query)

if not rsEmp.eof and not rsEmp.bof then 
%>
 <form name="frmEmpresa" id="frmEmpresa" action="solempresa/modificar.asp" method="post">
<table cellspacing="0" cellpadding="1" border=0>
	<tr>
    	<td width="100">Rut :</td>
      	<td width="166"><input id="txRut" readonly="readonly" name="txRut" type="text" tabindex="1" onKeyPress="" size="12" maxlength="11" value="<%=rsEmp("rut")%>"/></td>
        <td width="20" ><input type="hidden" id="txtId" name="txtId" value="<%=rsEmp("id_empresa")%>" /></td>
        <td width="100">&nbsp;</td>
        <td width="172">&nbsp;</td>
        <td width="20">&nbsp;</td>
        <td width="100">&nbsp;</td>
        <td width="131" >&nbsp;</td>
   	</tr>
	<tr>
        <td>Raz&oacute;n Social :</td>
        <td><input id="txtRsoc" name="txtRsoc" readonly="readonly" type="text" tabindex="2" maxlength="150" value="<%=rsEmp("r_social")%>" size="30"/></td>
        <td></td>
    	<td>Giro:</td>
        <td colspan="4"><input id="txtGiro" name="txtGiro" readonly="readonly" type="text" tabindex="3" maxlength="150" size="86" value="<%=rsEmp("giro")%>"/></td>
    </tr>
	<tr>
        <td>Direcci&oacute;n :</td>
        <td><input id="txtDir" name="txtDir" readonly="readonly" type="text" tabindex="4" maxlength="150" value="<%=rsEmp("direccion")%>" size="30"/></td>
        <td></td>
        <td>Comuna :</td>
        <td><input id="txtCom" name="txtCom" readonly="readonly" type="text" tabindex="5" size="30" maxlength="40" value="<%=rsEmp("comuna")%>" /></td>
        <td></td>
        <td>Ciudad :</td>
        <td><input id="txtCiu" name="txtCiu" readonly="readonly" type="text" tabindex="6"  size="30"  maxlength="30" value="<%=rsEmp("ciudad")%>" /></td>
    </tr>
	<tr>
        <td>Tel&eacute;fono :</td>
        <td><input id="txtFon" name="txtFon" readonly="readonly" type="text" tabindex="7" size="12" maxlength="50" value="<%=rsEmp("fono")%>"/></td>
        <td></td>
        <td>Fax :</td>
        <td><input id="txtFax" name="txtFax" readonly="readonly" type="text" tabindex="8" maxlength="50" value="<%=rsEmp("fax")%>" size="12"/></td>
	    <td></td>
    	<td>Mutual :</td>
        <td><select id="txtMut" name="txtMut" tabindex="9" disabled="disabled"></select><input id="txtIdMut" name="txtIdMut" type="hidden" value="<%=rsEmp("mutual")%>"/><input id="txtIdOtic" name="txtIdOtic" type="hidden" value="<%=rsEmp("ID_OTIC")%>"/></td>
    </tr>
    <tr>
    	<td colspan="8">&nbsp;<input id="tipoContacto" name="tipoContacto" type="hidden" value="<%=rsEmp("TIPO_CONTACTO")%>"/></td>
    </tr>
     <tr>
    	<td colspan="8"><center>
    	  <h3><em style="text-transform: capitalize;">Contacto de Cordinaci&oacute;n de Cursos</em></h3></center></td>
   	</tr>
     <tr>
    	<td colspan="8">&nbsp;</td>
    </tr>
    <tr>
        <td>Nombre :</td>
        <td><input id="txtNomb" name="txtNomb" readonly="readonly" type="text" tabindex="12" maxlength="50" value="<%=rsEmp("nombres")%>"  size="30" /></td>
        <td>&nbsp;</td>
    	<td>Cargo :</td>
        <td><input id="txtCargo" name="txtCargo" type="text" tabindex="15" readonly="readonly" maxlength="50" value="<%=rsEmp("cargo")%>" size="30"/></td>
        <td>&nbsp;</td>
        <td>T&eacute;lefono :</td>
        <td><input id="txtFonoCont" name="txtFonoCont" type="text" readonly="readonly" tabindex="14" maxlength="50" size="12" value="<%=rsEmp("FONO_CONTACTO")%>"/></td>        
    </tr>
     <tr>
        <td>Email :</td>
        <td colspan="7"><input id="txtMail" name="txtMail" type="text" readonly="readonly" tabindex="13" maxlength="50" value="<%=rsEmp("email")%>" size="30"/></td>
    </tr>
     <tr>
    	<td colspan="8">&nbsp;</td>
    </tr>
     <tr>
    	<td colspan="8"><center>
   	   <h3><em style="text-transform: capitalize;">Contacto de Contabilidad</em></h3></center></td>
   	</tr>
     <tr>
    	<td colspan="8">&nbsp;</td>
    </tr>
     <tr>
        <td>Nombre :</td>
        <td><input id="txtNombConta" name="txtNombConta" readonly="readonly" type="text" tabindex="17" maxlength="50" size="30" value="<%=rsEmp("NOMBRE_CONTA")%>"/></td>
        <td>&nbsp;</td>        
        <td>Cargo :</td>
        <td><input id="txtCargoConta" name="txtCargoConta" type="text" tabindex="20" readonly="readonly" maxlength="50" size="30" value="<%=rsEmp("CARGO_CONTA")%>"/></td>
        <td>&nbsp;</td>
        <td>T&eacute;lefono :</td>
        <td><input id="txtContaFono" name="txtContaFono" type="text" readonly="readonly" tabindex="19" maxlength="50" size="12" value="<%=rsEmp("FONO_CONTABILIDAD")%>"/></td>     
    </tr>
    <tr>
        <td>Email :</td>
        <td colspan="7"><input id="txtMailConta" name="txtMailConta" readonly="readonly" type="text" tabindex="18" maxlength="50" size="30" value="<%=rsEmp("EMAIL_CONTA")%>"/></td>
    </tr>
    <tr>
        <td colspan="8">&nbsp;</td>
    </tr>
     <tr>
        <td colspan="8"><input type="checkbox" name="rechazar" id="rechazar" onclick="estado_solicitud();" tabindex="21"/>Rechazar Solicitud de Nueva Empresa</td>
    </tr>
    <tr>
        <td colspan="8"><input id="rechazo" name="rechazo" type="hidden" value="0"/></td>
    </tr>
</table>
</form> 
<%
   end if
%>
<div id="messageBox1" style="height:40px;overflow:auto;width:400px;"> 
  	<ul></ul> 
</div> 
