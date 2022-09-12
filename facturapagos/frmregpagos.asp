<%
on error resume next
%>
<!--#include file="../conexion.asp"-->
<%
Response.CodePage = 65001
Response.CharSet = "utf-8"
vid = Request("id")

dim query
query= "select E.RUT,E.R_SOCIAL,FC.N_DOCUMENTO as ORDEN_COMPRA, "
query= query&"FC.FACTURA as FACTURA,FC.ID_EMPRESA as id_empresa, "
query= query&"CONVERT(VARCHAR(10),FC.FECHA_EMISION, 105) as FECHA_EMISION, " 
query= query&"CONVERT(VARCHAR(10),FC.FECHA_VENCIMIENTO, 105) as FECHA_VENCIMIENTO,FC.MONTO as VALOR_OC, "
query= query&"FC.DOCUMENTO_COMPROMISO,FC.ID_FACTURA from FACTURAS FC "
query= query&" inner join EMPRESAS E on E.ID_EMPRESA=FC.ID_EMPRESA " 
query= query&" WHERE FC.ID_FACTURA="&vid

set rsEmp = conn.execute (query)

if not rsEmp.eof and not rsEmp.bof then 
%>
 <form name="frmregpagos" id="frmregpagos" action="facturapagos/insertar.asp" method="post">
<table cellspacing="0" cellpadding="2" border=0>
	<tr>
    	<td>Rut :</td>
      	<td><%=rsEmp("RUT")%></td>
        <td><input type="hidden" id="Empresa" name="Empresa" value="<%=rsEmp("id_empresa")%>"/></td>
        <td>Raz&oacute;n Social :</td>
        <td colspan="4"><%=rsEmp("R_SOCIAL")%></td>
   	</tr>
	<tr>
        <td width="110">Factura :</td>
        <td width="160"><%=rsEmp("FACTURA")%></td>
        <td width="20">&nbsp;</td>
        <td width="120">Fecha Emisi&oacute;n :</td>
        <td width="145"><%=rsEmp("FECHA_EMISION")%></td>
        <td width="20">&nbsp;</td>
        <td width="128">Fecha Vencimiento :</td>
        <td width="130"><%=rsEmp("FECHA_VENCIMIENTO")%></td>
    </tr>
     <tr>
        <td>Total Factura :</td>
        <td><%="$ "&replace(FormatNumber(rsEmp("VALOR_OC"),0),",",".")%></td>
        <td><input type="hidden" id="txFolio" name="txFolio" value="<%=rsEmp("FACTURA")%>"/></td>
        <td>Saldo Pendiente :</td>
        <td><label id="lbSaldo"></label></td>
        <td>&nbsp;</td>
        <td>Monto a Cancelar :</td>
        <td><input id="txtMonto" name="txtMonto" type="text" tabindex="1" size="20" maxlength="10" onkeypress="return acceptNum(event)" onkeyup="calculaSaldo();"/></td>
    </tr>
    <tr>
        <td>Forma de Pago :</td>
        <td><select id="FormaPago" name="FormaPago" tabindex="5"></select></td>
   		<td><input type="hidden" id="idFact" name="idFact" value="<%=rsEmp("ID_FACTURA")%>"/></td>
        <td>N&deg; de Documento :</td>
        <td><input id="txDoc" name="txDoc" type="text" tabindex="3" size="20" maxlength="20" onkeypress="return acceptNum(event)"/></td>
        <td>&nbsp;</td>
        <%
		totalPago=0
		montoTotal=0
		pagos = "select count(INGRESOS.monto)as pagos from INGRESOS where INGRESOS.id_factura='"&rsEmp("ID_FACTURA")&"'"
		set rsPagos = conn.execute (pagos)

		if(rsPagos("pagos")<>"0")then
			abono = "select sum(INGRESOS.monto) as abono from INGRESOS where INGRESOS.id_factura='"&rsEmp("ID_FACTURA")&"'"
			set rsAbono = conn.execute (abono)
			totalPago=rsAbono("abono")
		end if
		montoTotal=cdbl(cdbl(rsEmp("VALOR_OC")) - cdbl(totalPago))
		%>
        <td>Fecha de Pago :
        <input type="hidden" id="Saldo" name="Saldo" value="<%=montoTotal%>"/></td>
        <td><input id="txtFechIngreso" name="txtFechIngreso" type="text" tabindex="4"  size="20"  maxlength="30"/></td>
    </tr>
    <tr>
        <td colspan="8">&nbsp;</td>
     </tr>
</table>
</form> 
<center><div style="width:800px;">
<table id="list2" align="left"></table> 
   	<div id="pager2"></div> 
</div></center>
<%
   end if
%>
<div id="messageBox1" style="height:70px;overflow:auto;width:300px;"> 
  	<ul></ul> 
</div> 
