<%
Response.CodePage = 65001
Response.CharSet = "utf-8"
%>

<form name="frmDelNota" id="frmDelNota" action="libroventas/delNota.asp" method="post">
   <table cellspacing="0" cellpadding="1" border=0>
    <tr>
       <td width="500">Seleccione que acción desea realizar con la Nota de Venta.</td>
    </tr>
    <tr><td>&nbsp;</td></tr>
    <tr>
    <td>Acción: &nbsp;<select id="tipo_bajaNota" name="tipo_bajaNota" tabindex="3">     
            <option value="" selected="selected">Seleccione</option>
            <option value="1">Solo Eliminar Nota de Venta</option>
            <option value="2">Eliminar Nota de Venta y Regresar a Facturación</option>
        </select>
       <input id="idFactNota" name="idFactNota" type="hidden" value="<%=Request("id_f")%>"/>
       <input id="idAutoNota" name="idAutoNota" type="hidden" value="<%=Request("id_a")%>"/>
       <input id="idUSRNota" name="idUSRNota" type="hidden" value="<%=Request("id_u")%>"/>
       </td>
    </tr>
    <tr><td>&nbsp;</td></tr>    
    <tr>
       <td>Razones : </td>
    </tr>
    <tr>
       <td><textarea name="razonDelNota" id="razonDelNota" rows="5" cols="70"></textarea></td>
    </tr>
    <tr>
       <td>&nbsp;</td>
    </tr>
   </table>
</form> 
<div id="messageBox2" style="height:70px;overflow:auto;width:400px;"> 
  	<ul></ul> 
</div> 