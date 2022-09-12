<!--#include file="conexion.asp"-->
<%
Response.CodePage = 65001
Response.CharSet = "utf-8"
vid = Request("id")
dim query
query= "select * from otic where id_otic="&vid
set rsEmp = conn.execute (query)

if not rsEmp.eof and not rsEmp.bof then 
%>
 <form name="frmOtic" id="frmOtic" action="Otic/modificar.asp" method="post">
<table cellspacing="0" cellpadding="1" border=0>
	<tr>
    	<td width="102">Rut :</td>
      	<td width="166"><input id="txRut" name="txRut" type="text" tabindex="1" onKeyPress="" size="12" maxlength="11" value="<%=rsEmp("rut")%>"/></td>
        <td width="20" ><input type="hidden" id="txtId" name="txtId" value="<%=rsEmp("id_otic")%>" /></td>
        <td width="126">&nbsp;</td>
        <td width="172">&nbsp;</td>
        <td width="20">&nbsp;</td>
        <td width="128">&nbsp;</td>
        <td width="131" >&nbsp;</td>
   	</tr>
	<tr>
        <td>Raz&oacute;n Social :</td>
        <td><input id="txtRsoc" name="txtRsoc" type="text" tabindex="3" maxlength="50" value="<%=rsEmp("r_social")%>" size="20"/></td>
        <td></td>
    	<td>Giro:</td>
        <td><input id="txtGiro" name="txtGiro" type="text" tabindex="4" size="20" maxlength="50"  value="<%=rsEmp("giro")%>"/></td>
    </tr>
	<tr>
        <td>Direcci&oacute;n :</td>
        <td><input id="txtDir" name="txtDir" type="text" tabindex="5" maxlength="20" value="<%=rsEmp("direccion")%>" size="20"/></td>
        <td></td>
        <td>Comuna :</td>
        <td><input id="txtCom" name="txtCom" type="text" tabindex="6" size="20" maxlength="40" value="<%=rsEmp("comuna")%>" /></td>
        <td></td>
        <td>Ciudad :</td>
        <td><input id="txtCiu" name="txtCiu" type="text" tabindex="7"  size="20"  maxlength="30" value="<%=rsEmp("ciudad")%>" /></td>
    </tr>
	<tr>
        <td>Tel&eacute;fono :</td>
        <td><input id="txtFon" name="txtFon" type="text" tabindex="8" size="12" maxlength="10"  value="<%=rsEmp("fono")%>" onKeyPress=""/></td>
        <td></td>
        <td>Fax :</td>
        <td><input id="txtFax" name="txtFax" type="text" tabindex="9" maxlength="10" value="<%=rsEmp("fax")%>"  size="12" onKeyPress=""/></td>
    </tr>
	<tr>
    	<td>Mutual :</td>
        <td><input id="txtMut" name="txtMut" type="text" tabindex="10" size="12" maxlength="11"  onkeyup="" value="<%=rsEmp("mutual")%>"/></td>
        <td></td>
        <td>Estable :</td>
        <td><input id="txtEst" name="txtEst" type="text" tabindex="11" maxlength="50" value="<%=rsEmp("estable")%>"  size="12"/></td>
        <td></td>
    	<td>Proyecto :</td>
        <td><input id="txtProy" name="txtProy" type="text" tabindex="12" maxlength="50" value="<%=rsEmp("proyecto")%>"  size="20"/></td>
    </tr>
    <tr>
    	<td>&nbsp;</td>
        <td>&nbsp;</td>
        <td></td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
    </tr>
    <tr>
    	<td>Rut Contacto :</td>
        <td><input id="txtRut2" name="txtRut2" type="text" tabindex="14" onKeyPress="" size="12" maxlength="12"  value="<%=rsEmp("rut_contacto")%>"/></td>
        <td></td>
        <td></td>
        <td></td>
    </tr>
     <tr>
        <td>Nombres :</td>
        <td><input id="txtNomb" name="txtNomb" type="text" tabindex="15" maxlength="50" value="<%=rsEmp("nombres")%>"  size="20"/></td>
        <td></td>
        <td>Apellido Paterno :</td>
        <td><input id="txtAp_pater" name="txtAp_pater" type="text" tabindex="16" maxlength="50" value="<%=rsEmp("a_paterno")%>"  size="20"/></td>
        <td></td>
        <td>Apellido Materno :</td>
        <td><input id="txtAp_mater" name="txtAp_mater" type="text" tabindex="17" maxlength="50" value="<%=rsEmp("a_materno")%>"  size="20"/></td>
    </tr>
     <tr>
    	<td>Celular :</td>
        <td><input id="txtCel" name="txtCel" type="text" tabindex="19" maxlength="50" value="<%=rsEmp("celular")%>"  size="12"/></td>
        <td></td>
        <td>Email :</td>
        <td><input id="txtMail" name="txtMail" type="text" tabindex="20" maxlength="50" value="<%=rsEmp("email")%>"  size="20"/></td>
    </tr>
	<tr>
    	<td>Cargo :</td>
        <td><input id="txtCargo" name="txtCargo" type="text" tabindex="21" maxlength="50" value="<%=rsEmp("cargo")%>"  size="12"/></td>
        <td></td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
    </tr>
</table>
</form> 
<%
   else
%><form name="frmOtic" id="frmOtic" action="Otic/insertar.asp" method="post">
	<table cellspacing="0" cellpadding="1" border=0 >
	<tr>
    	<td width="102">Rut :</td>
      	<td width="166"><input id="txRut" name="txRut" type="text" tabindex="1" size="12" maxlength="11" value=""/></td>
        <td width="20">&nbsp;</td>
        <td width="126">&nbsp;</td>
        <td width="172">&nbsp;</td>
        <td width="20">&nbsp;</td>
        <td width="128">&nbsp;</td>
        <td width="131">&nbsp;</td> 
   	</tr>
	<tr>
        <td>Raz&oacute;n Social :</td>
        <td><input id="txtRsoc" name="txtRsoc" type="text" tabindex="3" maxlength="50" value="" size="20"/></td>
        <td> </td>
    	<td>Giro:</td>
        <td><input id="txtGiro" name="txtGiro" type="text" tabindex="4" size="20" maxlength="50"  value=""/></td>
    </tr>
	<tr>
        <td>Direcci&oacute;n :</td>
        <td><input id="txtDir" name="txtDir" type="text" tabindex="5" maxlength="20" value="" size="20"/></td>
        <td></td>
        <td>Comuna :</td>
        <td><input id="txtCom" name="txtCom" type="text" tabindex="6" size="20" maxlength="40" value="" /></td>
        <td></td>
        <td>Ciudad :</td>
        <td><input id="txtCiu" name="txtCiu" type="text" tabindex="7"  size="20"  maxlength="30" value="" /></td>
    </tr>
	<tr>
        <td>Tel&eacute;fono :</td>
        <td><input id="txtFon" name="txtFon" type="text" tabindex="8" size="12" maxlength="10"  value=""/></td>
        <td></td>
        <td>Fax :</td>
        <td><input id="txtFax" name="txtFax" type="text" tabindex="9" maxlength="10" value=""  size="12" /></td>
    </tr>
	<tr>
    	<td>Mutual :</td>
        <td><input id="txtMut" name="txtMut" type="text" tabindex="10" size="12" maxlength="11"  onkeyup="" value=""/></td>
        <td></td>
        <td>Estable :</td>
        <td><input id="txtEst" name="txtEst" type="text" tabindex="11" maxlength="50" value=""  size="12"/></td>
        <td></td>
    	<td>Proyecto :</td>
        <td><input id="txtProy" name="txtProy" type="text" tabindex="12" maxlength="50" value=""  size="20"/></td>
    </tr>
    <tr>
    	<td>&nbsp;</td>
        <td>&nbsp;</td>
        <td></td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
    </tr>
    <tr>
    	<td>Rut Contacto :</td>
        <td><input id="txtRut2" name="txtRut2" type="text" tabindex="14" size="12" maxlength="12"  value=""/></td>
        <td></td>
        <td></td>
        <td></td>
    </tr>
     <tr>
        <td>Nombres :</td>
        <td><input id="txtNomb" name="txtNomb" type="text" tabindex="15" maxlength="50" value=""  size="20"/></td>
        <td></td>
        <td>Apellido Paterno :</td>
        <td><input id="txtAp_pater" name="txtAp_pater" type="text" tabindex="16" maxlength="50" value=""  size="20"/></td>
        <td></td>
        <td>Apellido Materno :</td>
        <td><input id="txtAp_mater" name="txtAp_mater" type="text" tabindex="17" maxlength="50" value=""  size="20"/></td>
    </tr>
     <tr>
    	<td>Celular :</td>
        <td><input id="txtCel" name="txtCel" type="text" tabindex="19" maxlength="50" value=""  size="12" /></td>
        <td></td>
        <td>Email :</td>
        <td><input id="txtMail" name="txtMail" type="text" tabindex="20" maxlength="50" value=""  size="20"/></td>
    </tr>
	<tr>
    	<td>Cargo :</td>
        <td><input id="txtCargo" name="txtCargo" type="text" tabindex="21" maxlength="50" value=""  size="12"/></td>
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
