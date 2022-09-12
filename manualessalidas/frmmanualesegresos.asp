<%
on error resume next
%>
<!--#include file="../conexion.asp"-->
<%
Response.CodePage = 65001
Response.CharSet = "utf-8"

dim query
query = "select A.DESC_ARTICULO AS ARTICULO_MOV,(B.UBICACION+' '+B.DIRECCION) AS BODEGA_MOV, "
query = query&"M.CANTIDAD,CONVERT(VARCHAR(10),P.FECHA_INICIO_, 105) as FECHA_INICIO, "
query = query&"DBO.MayMinTexto(IR.NOMBRES+' '+IR.A_PATERNO+' '+IR.A_MATERNO) AS RELATOR, "
query = query&"CONVERT(VARCHAR(10),M.FECHA, 105) as FECHA_MOV from MOVIMIENTOS M "
query = query&" INNER JOIN ARTICULOS A ON A.ID_ARTICULO=M.ID_ARTICULO "
query = query&" INNER JOIN BODEGAS B ON B.ID_BODEGA=M.ID_BODEGA "
query = query&" INNER JOIN bloque_programacion BP ON BP.id_bloque=M.ID_PROG_BLOQUE "
query = query&" INNER JOIN PROGRAMA P ON P.ID_PROGRAMA=BP.id_programa "
query = query&" INNER JOIN INSTRUCTOR_RELATOR IR ON IR.ID_INSTRUCTOR=BP.id_relator "
query = query&" where M.ID_MOVIMIENTO='"&Request("id")&"'"

set rsMov = conn.execute (query)

if not rsMov.eof and not rsMov.bof then 
%> 
<form name="frmmanualesegresos" id="frmmanualesegresos" action="#" method="post">
<table cellspacing="1" cellpadding="3" border=0>
    <tr>
    	<td width="130">Bodega :</td>
      	<td width="380"><%=rsMov("BODEGA_MOV")%></td>
        <td width="80" align="right">Fecha :</td>
        <td width="100"><%=rsMov("FECHA_MOV")%></td>
   	</tr>
    <tr>
    	<td>Articulo :</td>
   	    <td colspan="3"><%=rsMov("ARTICULO_MOV")%></td>
   	</tr>
    <tr>
    	<td>Cantidad :</td>
      	<td colspan="3"><%=rsMov("CANTIDAD")%></td>
   	</tr>
    <tr>
      <td>Fecha del Curso :</td>
      <td colspan="3"><%=rsMov("FECHA_INICIO")%></td>
    </tr>
    <tr>
      <td>Relator del Curso :</td>
      <td colspan="3"><%=rsMov("RELATOR")%></td>
    </tr>
</table>
</form> 
<%
   else
%> 
<form name="frmmanualesegresos" id="frmmanualesegresos" action="manualessalidas/insertar.asp" method="post">
<table cellspacing="0" cellpadding="3" border=0>
    <tr>
    	<td width="130">Bodega :</td>
      	<td width="200"><select id="MovBdg" name="MovBdg" tabindex="1" style="width:25em;" onchange="llena_articulos(0,this.value);"></select></td>
      <td width="80" align="right">Fecha :</td>
      <td width="80"><input id="MovFecha" name="MovFecha" type="text" tabindex="2" maxlength="50" size="12" value="<%=right("0"&day(now()),2)&"-"&right("0"&month(now()),2)&"-"&year(now)%>"/></td>
   	</tr>
    <tr>
    	<td>Articulo :</td>
   	    <td colspan="3"><select id="MovArt" name="MovArt" tabindex="3" style="width:42em;"></select></td>
   	</tr>
    <tr>
    	<td>Cantidad :</td>
      	<td colspan="3"><input id="MovCantidad" name="MovCantidad" type="text" tabindex="5" size="10" maxlength="10"/></td>
   	</tr>
    <tr>
      <td>Curso :</td>
      <td colspan="3"><select id="Curriculo" name="Curriculo" tabindex="6" onchange="llena_programa(this.value);" style="width:42em;"></select></td>
    </tr>    
    <tr>
      <td>Fecha del Curso :</td>
      <td colspan="3"><select id="ProgBuscar" name="ProgBuscar" tabindex="7" onchange="llena_bloque(this.value);" style="width:9em;"></select></td>
    </tr>
    <tr>
      <td>Relator del Curso :</td>
      <td colspan="3"><select id="MovBloque" name="MovBloque" tabindex="8" onchange=" $('#MovCodCurso').val(this.value);" style="width:25em;"></select><input id="MovCodCurso" name="MovCodCurso" type="hidden"/></td>
    </tr>
</table>
</form> 
<div id="messageBox1" style="height:120px;overflow:auto;width:450px;"> 
  	<ul></ul> 
</div>
<%
   end if
%>