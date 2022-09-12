<%
on error resume next
%>
<!--#include file="../conexion.asp"-->
<%
Response.CodePage = 65001
Response.CharSet = "utf-8"

dim query
query = "select A.DESC_ARTICULO AS ARTICULO_MOV,(B.UBICACION+' '+B.DIRECCION) AS BODEGA_MOV, "
query = query&" M.CANTIDAD,CONVERT(VARCHAR(10),M.FECHA, 105) as FECHA_MOV,TP.NOMBRE_MOV,M.EXPLICACION "
query = query&" from MOVIMIENTOS M "
query = query&" INNER JOIN ARTICULOS A ON A.ID_ARTICULO=M.ID_ARTICULO "
query = query&" INNER JOIN BODEGAS B ON B.ID_BODEGA=M.ID_BODEGA "
query = query&" INNER JOIN TIPO_MOVIMIENTO TP ON TP.ID_TIPO_MOV=M.TIPO_MOVIMIENTO "
query = query&"where M.ID_MOVIMIENTO='"&Request("id")&"'"

set rsMov = conn.execute (query)

if not rsMov.eof and not rsMov.bof then 
%> 
<form name="frmmanualesajustes" id="frmmanualesajustes" action="manualesajustes/insertar.asp" method="post">
<table cellspacing="0" cellpadding="3" border=0>
    <tr>
    	<td width="130">Bodega :</td>
      	<td width="330"><%=rsMov("BODEGA_MOV")%></td>
      <td width="80" align="right">Fecha :</td>
      <td width="80"><%=rsMov("FECHA_MOV")%></td>
   	</tr>
    <tr>
    	<td>Articulo :</td>
   	    <td colspan="3"><%=rsMov("ARTICULO_MOV")%></td>
   	</tr>
	<tr>
        <td>Tipo :</td>
        <td colspan="3"><%=rsMov("NOMBRE_MOV")%></td>
    </tr>    
    <tr>
    	<td>Cantidad :</td>
      	<td colspan="3"><%=rsMov("CANTIDAD")%></td>
   	</tr>
   <tr>
        <td colspan="4">&nbsp;</td>
    </tr>
    <tr>
        <td colspan="4">Explicaci&oacute;n :</td>
    </tr>
    <tr>
        <td colspan="4"><%=rsMov("EXPLICACION")%></td>
    </tr>
</table>
</form> 
<%
   else
%> 
<form name="frmmanualesajustes" id="frmmanualesajustes" action="manualesajustes/insertar.asp" method="post">
<table cellspacing="0" cellpadding="3" border=0>
    <tr>
    	<td width="130">Bodega :</td>
      	<td width="200"><select id="MovBdg" name="MovBdg" tabindex="1" onchange="llena_articulos(0,this.value);" style="width:25em;"></select></td>
      <td width="80" align="right">Fecha :</td>
      <td width="80"><input id="MovFecha" name="MovFecha" type="text" tabindex="2" maxlength="50" size="12" value="<%=right("0"&day(now()),2)&"-"&right("0"&month(now()),2)&"-"&year(now)%>"/></td>
   	</tr>
    <tr>
    	<td>Articulo :</td>
   	    <td colspan="3"><select id="MovArt" name="MovArt" tabindex="3" style="width:25em;"></select></td>
   	</tr>
	<tr>
        <td>Tipo :</td>
        <td colspan="3"><select id="MovTipo" name="MovTipo" tabindex="4" onchange="estadoTipoAjuste(this.value);" style="width:25em;"></select><input id="MovTipoAjusteSel" name="MovTipoAjusteSel" type="hidden"/></td>
    </tr>    
    <tr>
    	<td>Cantidad :</td>
      	<td colspan="3"><input id="MovCantidad" name="MovCantidad" type="text" tabindex="6" size="10" maxlength="10"/></td>
   	</tr>
   <tr>
        <td colspan="4">&nbsp;</td>
    </tr>
    <tr>
        <td colspan="4">Explicaci&oacute;n :</td>
    </tr>
    <tr>
        <td colspan="4"><textarea type="text" rows="3" cols="69" id="MovExpli" name="MovExpli" tabindex="8"></textarea></td>
    </tr>
</table>
</form> 
<div id="messageBox1" style="height:120px;overflow:auto;width:450px;"> 
  	<ul></ul> 
</div> 
<%
   end if
%>