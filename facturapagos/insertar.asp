<%Response.CodePage = 65001
Response.CharSet = "utf-8"
Response.ContentType="text/xml"
Response.Write("<?xml version='1.0' encoding='UTF-8'?>")
Response.Write("<insertar>") %>
<!--#include file="../conexion.asp"-->
<%
on error resume next
vidempresa = Request("Empresa")
vfolio=Request("txFolio")
vdoc=Request("txDoc")
vfechingreso=Request("txtFechIngreso")
vmonto=Request("txtMonto")
vformapago=Request("FormaPago")
idfact=Request("idFact")

dim query
query = "insert into ingresos (empresa,n_factura,comprobante,fecha_pago,monto,forma_pago,estado,id_factura)"
query = query&" values('"&vidempresa&"','"&vfolio&"','"&vdoc&"',CONVERT(datetime,'"&vfechingreso&"',105),"
query = query&"'"&vmonto&"','"&vformapago&"',1,'"&idfact&"')"

conn.execute (query)

montoCancelado=cdbl(cdbl(Request("Saldo")) - cdbl(vmonto))

if(montoCancelado=0)then
conn.execute ("update FACTURAS set ESTADO_CANCELACION=1 where ID_FACTURA="&idfact)
end if

if err.number <> 0 then
	Response.Write("<commit>false</commit>")
else
	Response.Write("<commit>true</commit>")
end if
Response.Write("</insertar>") 
%>