<%
on error resume next
%>
<!--#include file="../conexion.asp"-->
<%
Response.CodePage = 65001
Response.CharSet = "utf-8"
vid = Request("id")
dim query
query= "select U.*,(select COUNT(*) from INSTRUCTOR_RELATOR IR where IR.RUT=U.RUT AND IR.ESTADO=1) AS RELCREADO from USUARIOS U where U.ID_USUARIO="&vid
set rsEmp = conn.execute (query)

if not rsEmp.eof and not rsEmp.bof then 
%>
 <form name="frmUsuario" id="frmUsuario" action="usuario/modificar.asp" method="post">
<table cellspacing="0" cellpadding="1" border=0>
	<tr>
    	<td width="110">Rut :</td>
      	<td width="200"><%=rsEmp("rut")%><input id="txtRut" name="txtRut" type="hidden" value="<%=rsEmp("rut")%>"/></td>
        <td width="20" ><input type="hidden" id="txtId" name="txtId" value="<%=rsEmp("ID_USUARIO")%>" /></td>
        <td width="150">&nbsp;</td>
        <td width="172">&nbsp;</td>
        <td width="20">&nbsp;</td>
        <td width="160">&nbsp;</td>
        <td width="131" >&nbsp;</td>
   	</tr>
     <tr>
        <td>Nombres :</td>
        <td><input id="txtNomb" name="txtNomb" type="text" tabindex="2" maxlength="50" value="<%=rsEmp("nombres")%>" size="20"/></td>
        <td></td>
        <td>Apellido Paterno :</td>
        <td><input id="txtAp_pater" name="txtAp_pater" type="text" tabindex="3" maxlength="50" value="<%=rsEmp("a_paterno")%>" size="20"/></td>
        <td></td>
        <td>Apellido Materno :</td>
        <td><input id="txtAp_mater" name="txtAp_mater" type="text" tabindex="4" maxlength="50" value="<%=rsEmp("a_materno")%>" size="20" /></td>
    </tr>
     <tr>
    	<td>Tel&eacute;fono :</td>
        <td><input id="txtCel" name="txtCel" type="text" tabindex="5" maxlength="50" value="<%=rsEmp("celular")%>" onKeyPress="return acceptNum(event)" size="12"/></td>
        <td></td>
        <td>Email :</td>
        <td><input id="txtEmail" name="txtEmail" type="text" tabindex="6" maxlength="50" value="<%=rsEmp("EMAIL")%>"  size="20"/></td>
    </tr>
      <tr>
    	<td>Cargo :</td>
        <td><input id="txtCargo" name="txtCargo" type="text" tabindex="7" maxlength="50" value="<%=rsEmp("CARGO_EMPRESA")%>" size="20" /></td>
    </tr>
    <tr>
    	<td>Contrase&ntilde;a :</td>
        <td><input id="txtPass" name="txtPass" type="password" tabindex="8" maxlength="50" value="<%=rsEmp("CONTRASEÑA")%>" size="20" /></td>
    </tr>
     <tr>
     <%
		dim selRel, selCrear
		
		selRel=""
		if(rsEmp("TIPO_USUARIO")="1")then
		selRel="checked='checked'"
		end if
		
		selCrear=""
		if(rsEmp("RELCREADO")<>"0")then
		selCrear="disabled='disabled'"
		end if
		
		
		%>
        <td colspan="2"><input type="checkbox" name="addRelator" id="addRelator" tabindex="8" value="1" <%=selRel%> <%=selCrear%>/>Crear a Usuario como Relator</label></td> 
    </tr>    
     <tr>
    	<td colspan="8">&nbsp;</td>
    </tr>
    </table>
    <table cellspacing="0" cellpadding="1" border=0>
     <tr>
    	<td colspan="6"><center><h3><em style="text-transform: capitalize;">Permisos Generales</em></h3></center></td>
   	</tr>
     <tr>
    	<td colspan="6">&nbsp;</td>
   	</tr>
     <tr>
        <td width="117">&nbsp;</td>
        <%
		dim perm1,perm2,perm3,perm4,perm5
		
		if(rsEmp("PERMISO1")="1")then
		perm1="checked='checked'"
		end if
		
		if(rsEmp("PERMISO2")="1")then
		perm2="checked='checked'"
		end if
		
		if(rsEmp("PERMISO3")="1")then
		perm3="checked='checked'"
		end if
		
		if(rsEmp("PERMISO4")="1")then
		perm4="checked='checked'"
		end if
		
		if(rsEmp("PERMISO5")="1")then
		perm5="checked='checked'"
		end if

		if(rsEmp("PERMISO6")="1")then
		perm6="checked='checked'"
		end if
		
		if(rsEmp("PERMISO7")="1")then
		perm7="checked='checked'"
		end if	
		
		if(rsEmp("PERMISO8")="1")then
		perm8="checked='checked'"
		end if
		
		if(rsEmp("PERMISO9")="1")then
		perm9="checked='checked'"
		end if
		
		if(rsEmp("PERMISO10")="1")then
		perm10="checked='checked'"
		end if
		
		if(rsEmp("PERMISO11")="1")then
		perm11="checked='checked'"
		end if	
		
		if(rsEmp("PERMISO12")="1")then
		perm12="checked='checked'"
		end if		
		
		if(rsEmp("PERMISO13")="1")then
		perm13="checked='checked'"
		end if	
		
		if(rsEmp("PERMISO14")="1")then
		perm14="checked='checked'"
		end if
		
		if(rsEmp("PERMISO15")="1")then
		perm15="checked='checked'"
		end if		
		%>
    	<td width="170"><input type="checkbox" name="p1" id="p1" value="1" <%=perm1%>>Administraci&oacute;n</td>
        <td width="170"><input type="checkbox" name="p2" id="p2" value="1" <%=perm2%>>Operaci&oacute;n</td>
        <td width="170"><input type="checkbox" name="p3" id="p3" value="1" <%=perm3%>>Manejo de Cursos</td>
        <td width="170"><input type="checkbox" name="p5" id="p5" value="1" <%=perm5%>>BHP</td>
        <td width="170"><input type="checkbox" name="p6" id="p6" value="1" <%=perm6%>>Consultas</td>
    </tr>
    <tr>
        <td colspan="6">&nbsp;</td>
    </tr>
    <tr>
    	<td colspan="6"><center>
    	  <h3><em style="text-transform: capitalize;">Permisos Pesta&ntilde;a Finanzas</em></h3></center></td>
   	</tr>
     <tr>
    	<td colspan="6">&nbsp;</td>
   	</tr>
    <tr>
        <td><input type="checkbox" name="p4" id="p4" value="1" <%=perm4%> style="display:none"></td>
        <td><input type="checkbox" name="p7" id="p7" value="1" <%=perm7%> onclick="Permisos();">Facturaci&oacute;n</td>
        <td><input type="checkbox" name="p8" id="p8" value="1" <%=perm8%> onclick="Permisos();">Registro de Pagos</td>
        <td><input type="checkbox" name="p9" id="p9" value="1" <%=perm9%> onclick="Permisos();">Cuenta Corriente</td>
        <td><input type="checkbox" name="p10" id="p10" value="1" <%=perm10%> onclick="Permisos();">Inf. de Vencimiento</td>
        <td><input type="checkbox" name="p11" id="p11" value="1" <%=perm11%> onclick="Permisos();">Empresas</td>
    </tr>
    <tr>
        <td colspan="6">&nbsp;</td>
    </tr>
    <tr>
    	<td colspan="6"><center>
    	  <h3><em style="text-transform: capitalize;">Permisos Pesta&ntilde;a Manuales</em></h3></center></td>
   	</tr>
     <tr>
    	<td colspan="6">&nbsp;</td>
   	</tr>
    <tr>
        <td><input type="checkbox" name="p12" id="p12" value="1" <%=perm12%> style="display:none"></td>
        <td>&nbsp;</td>
        <td><input type="checkbox" name="p13" id="p13" value="1" <%=perm13%> onclick="PermisosManuales();">Administraci&oacute;n</td>
        <td><input type="checkbox" name="p14" id="p14" value="1" <%=perm14%> onclick="PermisosManuales();">Ingresos</td>
        <td><input type="checkbox" name="p15" id="p15" value="1" <%=perm15%> onclick="PermisosManuales();">Salidas</td>
        <td>&nbsp;</td>
    </tr>
    <tr>
    	<td colspan="6">&nbsp;</td>
    </tr>
	</table>
</form> 
<%
   else
%> <form name="frmUsuario" id="frmUsuario" action="usuario/insertar.asp" method="post">
<table cellspacing="0" cellpadding="1" border=0>
	<tr>
    	<td width="110">Rut :</td>
      	<td width="200"><input id="txRut" name="txRut" type="text" tabindex="1" onKeyPress="" size="12" maxlength="11"/></td>
        <td width="20" >&nbsp;</td>
        <td width="150">&nbsp;</td>
        <td width="172">&nbsp;</td>
        <td width="20">&nbsp;</td>
        <td width="160">&nbsp;</td>
        <td width="131" >&nbsp;</td>
   	</tr>
     <tr>
        <td>Nombres :</td>
        <td><input id="txtNomb" name="txtNomb" type="text" tabindex="2" maxlength="50" size="20"/></td>
        <td></td>
        <td>Apellido Paterno :</td>
        <td><input id="txtAp_pater" name="txtAp_pater" type="text" tabindex="3" maxlength="50" size="20"/></td>
        <td></td>
        <td>Apellido Materno :</td>
        <td><input id="txtAp_mater" name="txtAp_mater" type="text" tabindex="4" maxlength="50" size="20"/></td>
    </tr>
     <tr>
    	<td>Tel&eacute;fono :</td>
        <td><input id="txtCel" name="txtCel" type="text" tabindex="5" maxlength="50" size="12"  onKeyPress="return acceptNum(event)"/></td>
        <td></td>
        <td>Email :</td>
        <td><input id="txtEmail" name="txtEmail" type="text" tabindex="6" maxlength="50" size="20"/></td>
    </tr>
      <tr>
    	<td>Cargo :</td>
        <td><input id="txtCargo" name="txtCargo" type="text" tabindex="7" maxlength="50" size="20"/></td>
    </tr>
     <tr>
    	<td>Contrase&ntilde;a :</td>
        <td><input id="txtPass" name="txtPass" type="password" tabindex="8" maxlength="50" value="<%=rsEmp("CARGO_EMPRESA")%>" size="20" /></td>
    </tr>
     <tr>
        <td colspan="2"><input type="checkbox" name="addRelator" id="addRelator" tabindex="8" value="1"/>Crear a Usuario como Relator</label></td> 
    </tr>       
     <tr>
    	<td colspan="8">&nbsp;</td>
    </tr>
    </table>
    <table cellspacing="0" cellpadding="1" border=0>
     <tr>
    	<td colspan="6"><center><h3><em style="text-transform: capitalize;">Permisos</em></h3></center></td>
   	</tr>
     <tr>
    	<td colspan="6">&nbsp;</td>
   	</tr>
     <tr>
        <td width="117">&nbsp;</td>
    	<td width="170"><input type="checkbox" name="p1" id="p1" value="1">Administraci&oacute;n</td>
        <td width="170"><input type="checkbox" name="p2" id="p2" value="1">Operaci&oacute;n</td>
        <td width="170"><input type="checkbox" name="p3" id="p3" value="1">Manejo de Cursos</td>
        <td width="170"><input type="checkbox" name="p5" id="p5" value="1">BHP</td>
        <td width="170"><input type="checkbox" name="p6" id="p6" value="1">Consultas</td>      
    </tr>
    <tr>
        <td colspan="6">&nbsp;</td>
    </tr>
    <tr>
    	<td colspan="6"><center>
    	  <h3><em style="text-transform: capitalize;">Permisos Pesta&ntilde;a Finanzas</em></h3></center></td>
   	</tr>
     <tr>
    	<td colspan="6">&nbsp;</td>
   	</tr>
    <tr>
        <td><input type="checkbox" name="p4" id="p4" value="1" style="display:none"></td>
        <td><input type="checkbox" name="p7" id="p7" value="1" onclick="Permisos();">Facturaci&oacute;n</td>
        <td><input type="checkbox" name="p8" id="p8" value="1" onclick="Permisos();">Registro de Pagos</td>
        <td><input type="checkbox" name="p9" id="p9" value="1" onclick="Permisos();">Cuenta Corriente</td>
        <td><input type="checkbox" name="p10" id="p10" value="1" onclick="Permisos();">Inf. de Vencimiento</td>
        <td><input type="checkbox" name="p11" id="p11" value="1" onclick="Permisos();">Empresas</td>
    </tr>
    <tr>
        <td colspan="6">&nbsp;</td>
    </tr>
    <tr>
    	<td colspan="6"><center>
    	  <h3><em style="text-transform: capitalize;">Permisos Pesta&ntilde;a Manuales</em></h3></center></td>
   	</tr>
     <tr>
    	<td colspan="6">&nbsp;</td>
   	</tr>
    <tr>
        <td><input type="checkbox" name="p12" id="p12" value="1" style="display:none"></td>
        <td>&nbsp;</td>
        <td><input type="checkbox" name="p13" id="p13" value="1" onclick="PermisosManuales();">Administraci&oacute;n</td>
        <td><input type="checkbox" name="p14" id="p14" value="1" onclick="PermisosManuales();">Ingresos</td>
        <td><input type="checkbox" name="p15" id="p15" value="1" onclick="PermisosManuales();">Salidas</td>
        <td>&nbsp;</td>
    </tr>
    <tr>
    	<td colspan="6">&nbsp;</td>
    </tr>
	</table>
</form> 
<%
   end if
%>
<div id="messageBox1" style="height:70px;overflow:auto;width:300px;"> 
  	<ul></ul> 
</div> 
