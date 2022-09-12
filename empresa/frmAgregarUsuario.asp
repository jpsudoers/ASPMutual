<!--#include file="../conexion.asp"-->
<%
Response.CodePage = 65001
Response.CharSet = "utf-8"


vid = Request("id")
vem = Request("idEmpresa")

if(Request("id")="")then
	vid=0
end if
if(Request("idEmpresa")="")then
	vem=0
end if

dim query
query = "select ID_EMPRESAS_USUARIOS,ID_EMPRESA,NOMBRE, CARGO, TELEFONO, EMAIL, CONTRASENA"
query = query&" from EMPRESAS_USUARIOS where ID_EMPRESAS_USUARIOS = "&vid 

set usuarioEmpresa = conn.execute(query)


if not usuarioEmpresa.eof and not usuarioEmpresa.bof then 
%>
<form name="frmAgregarUsuario" id="frmAgregarUsuario" action="empresa/modificarUsuario.asp?" method="post" >
 <table cellspacing="0" cellpadding="1" border=0>
        <tr>
    	<td colspan="8"><center>
   	    <h3><em style="text-transform: capitalize;">Contacto de Coordinacion y pagos de Cursos</em></h3>
   	   </center></td>
   	</tr>
    <tr>
	    <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td><input id="txtempre" name="txtempre" type="hidden" tabindex="10" value="<%=usuarioEmpresa("ID_EMPRESA")%>"/><td>
		<td><input id="txtusuario" name="txtusuario" type="hidden" tabindex="11" value="<%=usuarioEmpresa("ID_EMPRESAS_USUARIOS")%>"/><td>
        <td>&nbsp;</td>
    </tr>
     <tr>
        <td>Nombre :</td>
        <td><input id="txtNomb" name="txtNomb" type="text" tabindex="12" maxlength="50" size="30" value="<%=usuarioEmpresa("NOMBRE")%>"/></td>
        <td>&nbsp;</td>
	</tr>
     <tr>
	   <td>Email :</td>
        <td><input id="txtMail" name="txtMail" type="text" tabindex="13" maxlength="50" size="30" value="<%=usuarioEmpresa("EMAIL")%>" onchange="validarMail();"/></td>
        <td>&nbsp;</td>
	</tr>
     <tr>
        <td>T&eacute;lefono :</td>
        <td><input id="txtFonoCont" name="txtFonoCont" type="text" tabindex="14" maxlength="20" size="12" value="<%=usuarioEmpresa("TELEFONO")%>"/></td>
    </tr>
    <tr>
    	<td>Cargo :</td>
        <!-- <td><input id="txtCargo" name="txtCargo" type="text" tabindex="15" maxlength="50" size="30" value="<%=usuarioEmpresa("CARGO")%> 
"/></td> -->
 <td><select id="txtCargo" name="txtCargo" tabindex="15">
			<option>contacto de coordinacion</option>
			<option>contacto de contabilidad</option>
		  </select></td>
        <td></td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
    </tr>
        <tr>
    	<td>Contrase&ntilde;a :</td>
        <td><input id="txtPassCord" name="txtPassCord" type="password" tabindex="16" maxlength="50" size="20" value="<%=usuarioEmpresa("CONTRASENA")%>"/></td>
    </tr>
     <tr>
    	<td colspan="8">&nbsp;</td>
    </tr>
   </table>
   
</form> 
<%
   else
   if vem <> 0 then
%> <form name="frmAgregarUsuario" id="frmAgregarUsuario" action="empresa/insertarUsuario.asp" method="post">
   <table cellspacing="1" cellpadding="1" border=0>
        <tr>
    	<td colspan="8"><center>
   	    <h3><em style="text-transform: capitalize;">Contacto de Coordinacion y pagos de Cursos</em></h3>
   	   </center></td>
   	</tr>
    <tr>
	    <td>&nbsp;</td>
        <td>&nbsp;</td>
		<td><input id="txtEmpresa" name="txtEmpresa" type="hidden" value="<%=vem%>"/><td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
    </tr>
     <tr>
        <td>Nombre :</td>
        <td><input id="txtNomb" name="txtNomb" type="text" tabindex="12" maxlength="50" size="30" /></td>
        <td>&nbsp;</td>
	</tr>
     <tr>
        <td>Email :</td>
        <td><input id="txtMail" name="txtMail" type="text" tabindex="13" maxlength="50" size="30" onchange="validarMail();"/></td>
        <td>&nbsp;</td>
	</tr>
     <tr>
        <td>T&eacute;lefono :</td>
        <td><input id="txtFonoCont" name="txtFonoCont" type="text" tabindex="14" maxlength="20" size="12" /></td>
    </tr>
    <tr>
    	<td>Cargo :</td>
        <!-- <td><input id="txtCargo2" name="txtCargo2" type="text" tabindex="15" maxlength="50" size="30" /></td> -->
		  <td><select id="txtCargo" name="txtCargo" tabindex="15">
			<option>contacto de coordinacion</option>
			<option>contacto de contabilidad</option>
		  </select></td>
        <td></td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
    </tr>
        <tr>
    	<td>Contrase&ntilde;a :</td>
        <td><input id="txtPassCord" name="txtPassCord" type="password" tabindex="16" maxlength="50" size="20"/></td>
    </tr>
     <tr>
		
    	<td colspan="8">&nbsp;</td>
    </tr>
   </table>
  
</form> 

<%
else 
 rutempresa = request("rutEmpresa")
%>
<form name="frmAgregarUsuario" id="frmAgregarUsuario" action="empresa/insertarUsuarioNuevaEmpresa.asp" method="post">
   <table cellspacing="1" cellpadding="1" border=0>
        <tr>
    	<td colspan="8"><center>
   	    <h3><em style="text-transform: capitalize;">Contacto de Coordinacion y pagos de Cursos</em></h3>
   	   </center></td>
   	</tr>
    <tr>
	
	    <td>&nbsp;</td>
        <td>&nbsp;</td>
		<td><input id="txtrut" name="txtrut" type="hidden" tabindex="12" maxlength="50" size="30"  value="<%=rutempresa%>"/><td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
    </tr>
     <tr>
        <td>Nombre :</td>
        <td><input id="txtNomb" name="txtNomb" type="text" tabindex="12" maxlength="50" size="30" /></td>
        <td>&nbsp;</td>
	</tr>
     <tr>
        <td>Email :</td>
        <td><input id="txtMail" name="txtMail" type="text" tabindex="13" maxlength="50" size="30" onchange="validarMail();"/></td>
        <td>&nbsp;</td>
	</tr>
     <tr>
        <td>T&eacute;lefono :</td>
        <td><input id="txtFonoCont" name="txtFonoCont" type="text" tabindex="14" maxlength="20" size="12" /></td>
    </tr>
    <tr>
    	<td>Cargo :</td>
        <!--<td><input id="txtCargo3" name="txtCargo3" type="text" tabindex="15" maxlength="50" size="30" /></td> -->
		 <td><select id="txtCargo" name="txtCargo" tabindex="15">
			<option>contacto de coordinacion</option>
			<option>contacto de contabilidad</option>
		  </select></td>
        <td></td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
    </tr>
        <tr>
    	<td>Contrase&ntilde;a :</td>
        <td><input id="txtPassCord" name="txtPassCord" type="password" tabindex="16" maxlength="50" size="20"/></td>
    </tr>
     <tr>
		
    	<td colspan="8">&nbsp;</td>
    </tr>
   </table>
  
</form> 
<%
end if
end if 

%>
<div id="messageBox1" style="height:90px;overflow:auto;width:350px;"> 
  	<ul></ul> 
</div> 