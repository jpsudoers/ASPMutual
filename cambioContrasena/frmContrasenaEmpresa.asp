<!--#include file="../conexion.asp"-->
<%
on error resume next

Response.CodePage = 65001
Response.CharSet = "utf-8"

dim query
query= "select CONTRASENA as hola from EMPRESAS_USUARIOS where ID_EMPRESAS_USUARIOS= "&Session("IdUsuario")
set rsEmp = conn.execute(query)

if not rsEmp.eof and not rsEmp.bof then 
%>
 <form name="frmContrasena" id="frmContrasena" action="cambioContrasena/modificarEmpresa.asp" method="post">
<table cellspacing="0" cellpadding="1" border=0>
	<tr>
    	<td width="350">Contrase&ntilde;a Antigua :<input type="hidden" id="txtContrasena" name="txtContrasena" value="<%=rsEmp("hola")%>" /></td>
      	<td width="200"><input id="ConAntigua" name="ConAntigua" type="password" tabindex="1" size="30" maxlength="50"/></td>
   	</tr>
     <tr>
    	<td>Contrase&ntilde;a Nueva :</td>
        <td><input id="NuevaCont" name="NuevaCont" type="password" tabindex="2" maxlength="50" size="30"/></td>
     </tr>
     <tr>
        <td>Repetir Contrase&ntilde;a Nueva :</td>
        <td><input id="RepConNueva" name="RepConNueva" type="password" tabindex="3" maxlength="50" size="30"/></td>
    </tr>
    </table>
</form> 
<%
   end if
%>
<div id="messageBox1" style="height:75px;overflow:auto;width:400px;"> 
  	<ul></ul> 
</div> 
