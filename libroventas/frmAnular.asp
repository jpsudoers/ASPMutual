<%
Response.CodePage = 65001
Response.CharSet = "utf-8"
%>

<form name="frmFactAnular" id="frmFactAnular" action="libroventas/anular.asp" method="post">
   <table cellspacing="0" cellpadding="1" border=0>
    <tr>
       <td width="500">Seleccione que acción realizar con la Factura N&deg; <%=Request("nfactura")%>.</td>
    </tr>
    <tr><td>&nbsp;</td></tr>
    <tr>
    <td>Acción: &nbsp;<select id="tipo_baja" name="tipo_baja" tabindex="3">     
            <option value="" selected="selected">Seleccione</option>
            <option value="0">Anular</option>
            <option value="3">Refacturar</option>
        </select>
       <input id="idFactAnular" name="idFactAnular" type="hidden" value="<%=Request("id_factura")%>"/>
       <input id="idAutoAnular" name="idAutoAnular" type="hidden" value="<%=Request("id_autorizacion")%>"/>
       </td>
    </tr>
    <tr><td>&nbsp;</td></tr>    
    <tr>
       <td>Razones : </td>
    </tr>
    <tr>
       <td><textarea name="razonAnular" id="razonAnular" rows="5" cols="70"></textarea></td>
    </tr>
    <tr>
       <td>&nbsp;</td>
    </tr>
   </table>
</form> 
<div id="messageBox2" style="height:70px;overflow:auto;width:400px;"> 
  	<ul></ul> 
</div> 