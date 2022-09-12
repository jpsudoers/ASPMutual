<%
on error resume next

Response.CodePage = 65001
Response.CharSet = "utf-8"
%> 
<form name="frmmanualesmov" id="frmmanualesmov" action="manualesmov/insertar.asp" method="post">
<table cellspacing="0" cellpadding="3" border=0>
    <tr>
    	<td width="150">Bodega :</td>
      	<td width="200"><select id="MovBdg" name="MovBdg" tabindex="1" style="width:25em;" onchange="llena_articulos(0,this.value);"></select></td>
      <td width="60" align="right">Fecha :</td>
      <td width="80"><input id="MovFecha" name="MovFecha" type="text" tabindex="2" maxlength="50" size="12" value="<%=right("0"&day(now()),2)&"-"&right("0"&month(now()),2)&"-"&year(now)%>"/></td>
   	</tr>
    <tr>
    	<td>Articulo :</td>
   	  <td colspan="3"><select id="MovArt" name="MovArt" tabindex="3" style="width:25em;"></select></td>
   	</tr>
	<tr>
        <td>Tipo :</td>
        <td colspan="3"><select id="MovTipo" name="MovTipo" tabindex="4" style="width:25em;"></select></td>
    </tr>
    <tr>
    	<td>Documento :</td>
      	<td colspan="3"><select id="MovTipoDoc" name="MovTipoDoc" tabindex="5" style="width:15em;">
          		 <OPTION VALUE="">Seleccione</OPTION>
          		 <OPTION VALUE="1">Orden de Compra</OPTION>
   				 <OPTION VALUE="2">Gu&iacute;a de Despacho</OPTION>
   				 <OPTION VALUE="3">Factura</OPTION>
           </select></td>
   	</tr>
    <tr>
    	<td>Fecha del Documento :</td>    
        <td colspan="3"><input id="MovFechaDoc" name="MovFechaDoc" type="text" tabindex="2" maxlength="50" size="12"/></td>
    </tr>
    <tr>
    	<td>Cantidad :</td>
      	<td colspan="3"><input id="MovCantidad" name="MovCantidad" type="text" tabindex="6" size="10" maxlength="10"/></td>
   	</tr>
    <tr>
    	<td>Precio Unitario Con Iva :</td>
      	<td colspan="3"><input id="MovPrecio" name="MovPrecio" type="text" tabindex="7" size="10" maxlength="10"/> 
      	<font color="#FF0000"> * Ejemplo de Precios con Decimal: 100.78</font></td>
   	</tr>
</table>
</form> 
<div id="messageBox1" style="height:175px;overflow:auto;width:450px;"> 
  	<ul></ul> 
</div> 