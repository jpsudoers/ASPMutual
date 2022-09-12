<%
on error resume next
%>
<!--#include file="../conexion.asp"-->
<%
Response.CodePage = 65001
Response.CharSet = "utf-8"

dim query
query= "select * from BODEGAS where ID_BODEGA="&Request("id")
set rsBdg = conn.execute (query)

if not rsBdg.eof and not rsBdg.bof then 
%>
 <form name="frmbodegas" id="frmbodegas" action="bodegas/modificar.asp" method="post">
<table cellspacing="1" cellpadding="2" border=0>
    <tr>
    	<td width="110">C&oacute;digo :</td>
      	<td width="600"><%=right("000000000"&rsBdg("ID_BODEGA"),5)%></td>
   	</tr>
    <tr>
    	<td>Nombre :</td>
      	<td><input id="UbiBdg" name="UbiBdg" type="text" tabindex="1" size="50" maxlength="99" value="<%=rsBdg("UBICACION")%>"/></td>
   	</tr>
  	<tr>
        <td>Direcci&oacute;n :</td>
        <td><input id="DirBdg" name="DirBdg" type="text" tabindex="2" size="50" maxlength="99" value="<%=rsBdg("DIRECCION")%>"/></td>
    </tr>
     <tr>
         <td>Responsable :</td>
         <td><select id="Responsable" name="Responsable" tabindex="3" style="width:27em;"></select><input type="hidden" id="IdResp" name="IdResp" value="<%=rsBdg("RESPONSABLE")%>"/></td>
    </tr>
    <tr>
         <td colspan="2"><input type="hidden" id="IdBdg" name="IdBdg" value="<%=rsBdg("ID_BODEGA")%>"/></td>
    </tr>
</table>
</form> 
<%
   else
%> <form name="frmbodegas" id="frmbodegas" action="bodegas/insertar.asp" method="post">
<table cellspacing="1" cellpadding="2" border=0>
    <tr>
    	<td width="110">Nombre :</td>
      	<td width="600"><input id="UbiBdg" name="UbiBdg" type="text" tabindex="1" size="50" maxlength="99"/></td>
   	</tr>
    <tr>
        <td>Direcci&oacute;n :</td>
        <td><input id="DirBdg" name="DirBdg" type="text" tabindex="2" size="50" maxlength="99"/></td>
    </tr>
    <tr>
        <td>Responsable :</td>
         <td><select id="Responsable" name="Responsable" tabindex="3" style="width:27em;"></select></td>
    </tr>
    <tr>
         <td colspan="2">&nbsp;</td>
    </tr>
</table>
</form> 
<%
   end if
%>
<div id="messageBox1" style="height:100px;overflow:auto;width:300px;"> 
  	<ul></ul> 
</div> 
