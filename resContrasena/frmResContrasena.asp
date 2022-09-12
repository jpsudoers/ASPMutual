<%
on error resume next

Response.CodePage = 65001
Response.CharSet = "utf-8"

dim user
user=""
if(Request("usuario")<>"ID de Usuario")then
user=Request("usuario")
end if

dim empresa
empresa=""
if(Request("empresa")<>"Rut de Empresa")then
empresa=Request("empresa")
end if
%>
 <form name="frmResContrasena" id="frmResContrasena" action="resContrasena/EnResContrasena.asp" method="post">
<table cellspacing="0" cellpadding="1" border=0>
     <tr>
    	<td colspan="2" align="justify">Para poder restablecer la contrase&ntilde;a, debe escribir la informaci&oacute;n que a continuaci&oacute;n se le solicita.</td>
     </tr>
     <tr>
    	<td colspan="2">&nbsp;</td>
     </tr>
	<tr>
    	<td width="230">Rut de Empresa : </td>
      	<td width="300"><input id="resRutEmpresa" name="resRutEmpresa" type="text" tabindex="1" size="15" maxlength="12" value="<%=empresa%>"/></td>
   	</tr>
     <tr>
    	<td>&nbsp;</td>
        <td>Ejemplo: 16313883-7</td>
     </tr>
     <tr>
    	<td>&nbsp;</td>
        <td>&nbsp;</td>
     </tr>
     <tr>
    	<td>Correo del Usuario :</td>
        <td><input id="resEmailEmpresa" name="resEmailEmpresa" type="text" tabindex="2" maxlength="40" size="40" value="<%=user%>"/></td>
     </tr>
     <tr>
    	<td>&nbsp;</td>
        <td>Ejemplo: contacto@empresa.cl</td>
     </tr>
     <tr>
    	<td colspan="2">&nbsp;</td>
     </tr>
     <tr>
		<td colspan="2">&nbsp;</td>
     </tr>
     <tr>
    	<td colspan="2" align="justify"><b><i>El restablecimiento de su contrase&ntilde;a cambia autom&aacute;ticamente su antigua contrase&ntilde;a por una nueva contrase&ntilde;a, la cual ser&aacute; enviada en los pr&oacute;ximos minutos a su correo electr&oacute;nico.</i></b></td>
     </tr>
     <tr>
    	<td colspan="2">&nbsp;</td>
     </tr>
    </table>
</form> 
<div id="messageBox1" style="height:75px;overflow:auto;width:400px;"> 
  	<ul></ul> 
</div> 