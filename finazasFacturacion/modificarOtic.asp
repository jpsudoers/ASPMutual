<!--#include file="../conexion.asp"-->
<%
Response.CodePage = 65001
Response.CharSet = "utf-8"
Response.ContentType="text/xml"
Response.Write("<?xml version='1.0' encoding='UTF-8'?>")
Response.Write("<modificar>") 

on error resume next

idAuto = Request("idAutorizacionOtic")

valorempresa = Request("ValorEmpresa")
nfactura = Request("Nfactura")
idEmpresa = Request("idEmpresaOtic")

valotic = Request("ValOtic")
nfacturaotic = Request("NfacturaOtic")
idotic = Request("IdOtic")

ndocfact = Request("NDocFactOtic")
formapago = Request("FormaPagoOtic")
doc = Request("DOC_COMP_Otic")

dim query 
if(cdbl(valorempresa)>0)then
query = "insert into FACTURAS (ID_AUTORIZACION,FACTURA,FECHA_EMISION,FECHA_VENCIMIENTO,ESTADO_CANCELACION,FACTURADO,ID_EMPRESA,"
query = query&"MONTO,TIPO_DOC,DOCUMENTO_COMPROMISO,N_DOCUMENTO) values ('"&idAuto&"','"&nfactura&"',"
query = query&"GETDATE(),CONVERT(datetime, CONVERT(varchar,GETDATE()+30,105),105),'"&pagada&"','"&Session("usuarioMutual")&"',"
query = query&"'"&idEmpresa&"','"&valorempresa&"','"&formapago&"','"&doc&"','"&ndocfact&"')"

conn.execute (query)
end if

conn.execute ("update AUTORIZACION set FACTURADO=0 WHERE ID_AUTORIZACION='"&idAuto&"'")

dim queryOtic 
if(cdbl(valotic)>0)then
queryOtic = "insert into FACTURAS(ID_AUTORIZACION,FACTURA,FECHA_EMISION,FECHA_VENCIMIENTO,ESTADO_CANCELACION,FACTURADO,"
queryOtic = queryOtic&"ID_EMPRESA,MONTO,TIPO_DOC,DOCUMENTO_COMPROMISO,N_DOCUMENTO) values ('"&idAuto&"','"&nfacturaotic&"',"
queryOtic = queryOtic&"GETDATE(),CONVERT(datetime, CONVERT(varchar,GETDATE()+30,105),105),'"&pagada&"',"
queryOtic = queryOtic&"'"&Session("usuarioMutual")&"','"&idotic&"','"&valotic&"','"&formapago&"','"&doc&"','"&ndocfact&"')"

conn.execute (queryOtic)
end if

'set rs = conn.execute ("select IDENT_CURRENT('FACTURAS')AS UltFact")

'dim SQL
'if(pagada="1")then
	'SQL = "insert into ingresos (empresa,n_factura,comprobante,fecha_pago,monto,forma_pago,estado,id_factura)"
	'SQL = SQL&" values('"&idEmpresa&"','"&factura&"','"&ndocfact&"',CONVERT(datetime,'"&fechapago&"',105),'"&monto&"',"
	'SQL = SQL&"'"&formapago&"',1,'"&rs("UltFact")&"')"
	
	'conn.execute (SQL)
'end if

if err.number <> 0 then
	Response.Write("<commit>false</commit>")
else
	Response.Write("<commit>true</commit>")
end if
Response.Write("</modificar>") 
%>