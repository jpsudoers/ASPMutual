<%
on error resume next
%>
<!--#include file="../conexion.asp"-->
<%

Response.CodePage = 65001
Response.CharSet = "utf-8"
vid = Request("id")
dim query
query= "select EMPRESAS.ID_EMPRESA as empresa,EMPRESAS.RUT,EMPRESAS.R_SOCIAL,AUTORIZACION.ORDEN_COMPRA,AUTORIZACION.FACTURA, "
query= query&"AUTORIZACION.VALOR_OC,AUTORIZACION.ID_AUTORIZACION,AUTORIZACION.VALOR_OC"
query= query&" from AUTORIZACION "
query= query&" inner join EMPRESAS on EMPRESAS.ID_EMPRESA=AUTORIZACION.ID_EMPRESA "
query= query&" WHERE AUTORIZACION.ID_AUTORIZACION="&vid

set rsEmp = conn.execute (query)

if not rsEmp.eof and not rsEmp.bof then 
%>
 <form name="frmregfactura" id="frmregfactura" action="regfactura/modificar.asp" method="post">
<table cellspacing="0" cellpadding="1" border=0>
	<tr>
    	<td>Rut :</td>
      	<td><%=rsEmp("RUT")%></td>
        <td><input type="hidden" id="txtIdEmpresa" name="txtIdEmpresa" value="<%=rsEmp("empresa")%>"/>
        </td>
        <td>Raz&oacute;n Social :</td>
        <td colspan="4"><label id="txtRsoc" name="txtRsoc"><%=rsEmp("R_SOCIAL")%></label></td>
   	</tr>
	<tr>
        <td>Orden Compra :</td>
        <td><%=rsEmp("ORDEN_COMPRA")%></td>
        <td><input id="OrdenC" name="OrdenC" type="hidden" value="<%=rsEmp("ID_AUTORIZACION")%>"/></td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td></td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
    </tr>
    <tr>
   		<td width="102">N&deg; Factura :</td>
        <td width="166"><input id="txFolio" name="txFolio" type="text" tabindex="3" size="20" maxlength="20" onKeyPress="return acceptNum(event)"/></td>
        <td width="20">&nbsp;</td>
    	<td width="126">Fecha Emisi&oacute;n :</td>
        <td width="172"><input id="txtFechEmision" name="txtFechEmision" type="text" tabindex="4"  size="20"  maxlength="15"/></td>
        <td width="20"></td>
        <td width="128">Fecha Vencimiento :</td>
        <td width="131"><input id="txtFechVenc" name="txtFechVenc" type="text" tabindex="5"  size="20" maxlength="15"/></td>
    </tr>
</table>
</form> 
<%
   end if
%>
<div id="messageBox1" style="height:60px;overflow:auto;width:300px;"> 
  	<ul></ul> 
</div> 
