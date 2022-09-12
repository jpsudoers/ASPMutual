<%
on error resume next
%>
<!--#include file="../conexion.asp"-->
<%
Response.CodePage = 65001
Response.CharSet = "utf-8"

dim query
query= "SELECT ID_ARTICULO_BODEGA,ID_BODEGA,ID_ARTICULO_PROV,ART_MINIMOS,STOCK_CRITICO,REP_ARTICULO,STOCK_ACTUAL "
query= query&"FROM ARTICULO_BODEGA where ID_ARTICULO_BODEGA='"&Request("IdArtBod")&"'"

'response.Write(query)
'response.End()

set rsArt = conn.execute (query)

fecha = year(now)&right("0"&month(now()),2)&right("0"&day(now()),2)
fecha = fecha&right("0"&hour(now()),2)&right("0"&minute(now()),2)&right("0"&second(now()),2)

if not rsArt.eof and not rsArt.bof then 
%>
 <form name="frmasignar" id="frmasignar" action="manuales/modificar_Temp.asp" method="post">
<table cellspacing="0" cellpadding="3" border=0>
    <tr>
    	<td width="170">Bodega :</td>
      	<td width="330"><select id="BdgArt" name="BdgArt" tabindex="10" style="width:25em;"></select><input id="IdBdgSel" name="IdBdgSel" type="hidden" value="<%=rsArt("ID_BODEGA")%>"/></td>
   	</tr>
    <tr>
    	<td>Cantidad M&iacute;nima :</td>
      	<td><input id="CantMinArt" name="CantMinArt" type="text" tabindex="20" size="10" maxlength="10" value="<%=rsArt("ART_MINIMOS")%>"/></td>
   	</tr>
    <tr>
    	<td>Stock Critico : </td>
      	<td><input id="StockCritArt" name="StockCritArt" type="text" tabindex="30" size="10" maxlength="10" value="<%=rsArt("STOCK_CRITICO")%>"/></td>
   	</tr>
    <tr>
    	<td>Cantidad de Reposici&oacute;n :</td>
      	<td><input id="CantRepArt" name="CantRepArt" type="text" tabindex="40" size="10" maxlength="10" value="<%=rsArt("REP_ARTICULO")%>"/><input id="ID_Art_Bdg" name="ID_Art_Bdg" type="hidden" value="<%=rsArt("ID_ARTICULO_BODEGA")%>"/></td>
   	</tr>
    <tr>
    	<td>Stock Actual :</td>
      	<td><label id="StActual"><%=rsArt("STOCK_ACTUAL")%></label></td>
   	</tr>
</table>
</form> 
<%
   else
%> 
<form name="frmasignar" id="frmasignar" action="manuales/insertar_Temp.asp" method="post">
<table cellspacing="0" cellpadding="3" border=0>
    <tr>
    	<td width="170">Bodega :</td>
      	<td width="330"><select id="BdgArt" name="BdgArt" tabindex="10" style="width:25em;"></select><input id="IdBdgSel" name="IdBdgSel" type="hidden" value="0"/></td>
   	</tr>
    <tr>
    	<td>Cantidad M&iacute;nima :</td>
      	<td><input id="CantMinArt" name="CantMinArt" type="text" tabindex="20" size="10" maxlength="10"/></td>
   	</tr>
    <tr>
    	<td>Stock Critico : </td>
      	<td><input id="StockCritArt" name="StockCritArt" type="text" tabindex="30" size="10" maxlength="10"/></td>
   	</tr>
    <tr>
    	<td>Cantidad de Reposici&oacute;n :</td>
      	<td><input id="CantRepArt" name="CantRepArt" type="text" tabindex="40" size="10" maxlength="10"/><input id="IdProvAsig" name="IdProvAsig" type="hidden" value="<%=Request("IdProv")%>"/><input id="TipoInsert" name="TipoInsert" type="hidden" value="<%=Request("TipInsert")%>"/></td>
   	</tr>
</table>
</form> 
<%
   end if
%>
<div id="messageBox2" style="height:100px;overflow:auto;width:400px;"> 
  	<ul></ul> 
</div> 
