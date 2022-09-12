<%
on error resume next
%>
<!--#include file="../conexion.asp"-->
<%
Response.CodePage = 65001
Response.CharSet = "utf-8"

dim query
query= "select A.*,(SELECT COUNT(*) FROM ARTICULO_BODEGA AB WHERE AB.ID_ARTICULO=A.ID_ARTICULO) AS TOTAL_BODEGAS from ARTICULOS A where A.ID_ARTICULO='"&Request("id")&"'"

set rsArt = conn.execute (query)

fecha = year(now)&right("0"&month(now()),2)&right("0"&day(now()),2)
fecha = fecha&right("0"&hour(now()),2)&right("0"&minute(now()),2)&right("0"&second(now()),2)

if not rsArt.eof and not rsArt.bof then 
%>
 <form name="frmmanuales" id="frmmanuales" action="manuales/modificar.asp" method="post">
<table cellspacing="0" cellpadding="3" border=0>
    <tr>
    	<td>C&oacute;digo :</td>
      	<td><%=right("000000000"&rsArt("ID_ARTICULO"),5)%></td>
   	</tr>
    <tr>
    	<td width="170">C&oacute;digo Adicional :</td>
      	<td width="330"><input id="CodAdicArt" name="CodAdicArt" type="text" tabindex="1" size="40" maxlength="60" value="<%=rsArt("CODIGO_ADICIONAL")%>"/></td>
   	</tr>
  <tr>
        <td>Descripci&oacute;n :</td>
        <td><textarea type="text" rows="3" cols="50" id="DescArt" name="DescArt" tabindex="2"><%=rsArt("DESC_ARTICULO")%></textarea></td>
    </tr>
    <tr>
        <td>Unidad de Medida :</td>
      <td><select id="UMArt" name="UMArt" tabindex="3" style="width:27em;"></select><input id="UMArtID" name="UMArtID" type="hidden" value="<%=rsArt("UNIDAD_MEDIDA")%>"/></td>
    </tr>
    <tr>
        <td>Tipo de Articulo :</td>
        <td><select id="TipoArt" name="TipoArt" tabindex="4" style="width:27em;"></select><input id="TipoArtID" name="TipoArtID" type="hidden" value="<%=rsArt("ID_DPTO")%>"/></td>
    </tr>
    <tr>
      <td colspan="2">&nbsp;<input id="IdProvArt" name="IdProvArt" type="hidden" value="<%=rsArt("ID_ARTICULO")%>"/><input id="TotBdg" name="TotBdg" type="hidden" value="<%=rsArt("TOTAL_BODEGAS")%>"/><input id="EstadoInsert" name="EstadoInsert" type="hidden"/><input id="EstBdg" name="EstBdg" type="hidden" value="<%=fecha%>"/></td>
    </tr>
</table>
</form> 
<%
   else
%> <form name="frmmanuales" id="frmmanuales" action="manuales/insertar.asp" method="post">
<table cellspacing="0" cellpadding="3" border=0>
    <tr>
    	<td width="170">C&oacute;digo Adicional :</td>
      	<td width="330"><input id="CodAdicArt" name="CodAdicArt" type="text" tabindex="1" size="40" maxlength="60"/></td>
   	</tr>
  <tr>
        <td>Descripci&oacute;n :</td>
        <td><textarea type="text" rows="3" cols="50" id="DescArt" name="DescArt" tabindex="2"/></textarea></td>
    </tr>
    <tr>
        <td>Unidad de Medida :</td>
         <td><select id="UMArt" name="UMArt" tabindex="3" style="width:27em;"></select></td>
    </tr>
     <tr>
        <td>Tipo de Articulo :</td>
         <td><select id="TipoArt" name="TipoArt" tabindex="4" style="width:27em;"></select></td>
    </tr>
    <tr>
      <td colspan="2">&nbsp;<input id="IdProvArt" name="IdProvArt" type="hidden" value="<%=fecha%>"/><input id="TotBdg" name="TotBdg" type="hidden" value="0"/><input id="EstadoInsert" name="EstadoInsert" type="hidden"/><input id="EstBdg" name="EstBdg" type="hidden" value="<%=fecha%>"/></td>
    </tr>
</table>
</form> 
<%
   end if
%>
<div width="200"><table id="list2"></table> 
   		<div id="pager2"></div> 
</div> 
<div id="messageBox1" style="height:100px;overflow:auto;width:400px;"> 
  	<ul></ul> 
</div> 
