<%
on error resume next
%>
<!--#include file="../conexion.asp"-->
<%

Response.CodePage = 65001
Response.CharSet = "utf-8"
vid = Request("id")
dim query
query= "SELECT * from SEDES WHERE ID_SEDE="&vid
set rsEmp = conn.execute (query)

if not rsEmp.eof and not rsEmp.bof then 
%>
 <form name="frmSedes" id="frmSedes" action="Sedes/modificar.asp" method="post">
<table cellspacing="0" cellpadding="1" border=0>
	<tr>
    	<td width="102">Nombre :</td>
      <td width="166"><input id="txtNombre" name="txtNombre" type="text" tabindex="1" size="30" maxlength="50" value="<%=rsEmp("NOMBRE")%>"/></td>
        <td width="20" >&nbsp;</td>
        <td width="126">&nbsp;</td>
        <td width="172">&nbsp;</td>
        <td width="20">&nbsp;</td>
        <td width="128">&nbsp;</td>
        <td width="131" ><input type="hidden" id="txtId" name="txtId" value="<%=rsEmp("ID_SEDE")%>" /></td>
   	</tr>
	<tr>
        <td>Direcci&oacute;n :</td>
        <td><input id="txtDir" name="txtDir" type="text" tabindex="2" maxlength="50" value="<%=rsEmp("direccion")%>" size="30"/></td>
        <td>&nbsp;</td>
        <td>Comuna :</td>
        <td><input id="txtCom" name="txtCom" type="text" tabindex="3" size="20" maxlength="40" value="<%=rsEmp("comuna")%>" /></td>
        <td></td>
        <td>Ciudad :</td>
        <td><input id="txtCiu" name="txtCiu" type="text" tabindex="4" size="20" maxlength="30" value="<%=rsEmp("ciudad")%>" /></td>
    </tr>
    <tr>
    	<td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
    </tr>
 </table>
</form> 
<%
   else
%><form name="frmSedes" id="frmSedes" action="Sedes/insertar.asp" method="post">
	<table cellspacing="0" cellpadding="1" border=0 >
		<tr>
    	<td width="102">Nombre :</td>
      <td width="166"><input id="txtNombre" name="txtNombre" type="text" tabindex="1" size="30" maxlength="50"/></td>
        <td width="20" >&nbsp;</td>
        <td width="126">&nbsp;</td>
        <td width="172">&nbsp;</td>
        <td width="20">&nbsp;</td>
        <td width="128">&nbsp;</td>
        <td width="131" >&nbsp;</td>
   	</tr>
	<tr>
        <td>Direcci&oacute;n :</td>
        <td><input id="txtDir" name="txtDir" type="text" tabindex="2" maxlength="50" size="30"/></td>
        <td>&nbsp;</td>
        <td>Comuna :</td>
        <td><input id="txtCom" name="txtCom" type="text" tabindex="3" size="20" maxlength="40"/></td>
        <td></td>
        <td>Ciudad :</td>
        <td><input id="txtCiu" name="txtCiu" type="text" tabindex="4" size="20" maxlength="30"/></td>
    </tr>
    <tr>
    	<td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
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
