<!--#include file="../../conexion.asp"-->
<%
Response.CodePage = 65001
Response.CharSet = "utf-8"
Response.ContentType="text/xml"
Response.Write("<?xml version='1.0' encoding='UTF-8'?>")
Response.Write("<insertar>") 

on error resume next
idAuto=request("idAuto")
txtNFemp=request("txtNFemp") 
pagada=request("pagada")
usuario=request("usuario")
idEmpresa=request("idEmpresa") 
monto=request("monto")
formapago=request("formapago")
documento=request("documento")
ndocfact=request("ndocfact")
fechapago=request("fechapago")
fechaemision=request("fechaemision")

N_REG_SENCE="NULL"
if(request("N_REG_SENCE")<>"")then
	N_REG_SENCE="'"&request("N_REG_SENCE")&"'"
end if

query = "insert into FACTURAS (ID_AUTORIZACION,FACTURA,FECHA_EMISION,FECHA_VENCIMIENTO,ESTADO_CANCELACION,FACTURADO,"
query = query&"ID_EMPRESA,MONTO,TIPO_DOC,DOCUMENTO_COMPROMISO,N_DOCUMENTO,ESTADO) values"
query = query&"('" & idAuto & "','" & txtNFemp & "',convert(date,'"&fechaemision&"'),"
query = query&"CONVERT(datetime, CONVERT(varchar,convert(datetime,'"&fechaemision&"')+30,105),105),'"
query = query&pagada & "','" & usuario & "','" & idEmpresa & "','" & monto & "','" & formapago & "','"
query = query&documento & "','" & ndocfact &"',1)"

conn.execute (query)

conn.execute ("update AUTORIZACION set FACTURADO=0,N_REG_SENCE="&N_REG_SENCE&",ORDEN_COMPRA='"&ndocfact&"' WHERE ID_AUTORIZACION='"&idAuto&"'")

queryPagos=""
if (pagada="1") then

Set rsUlt = conn.execute("select IDENT_CURRENT('FACTURAS')AS UltFact")

queryPagos="insert into ingresos (empresa,n_factura,comprobante,fecha_pago,monto,forma_pago,estado,id_factura)"
queryPagos=queryPagos&" values('"&idEmpresa&"','"&txtNFemp&"','"&ndocfact&"',CONVERT(datetime,'"&fechapago&"',105),'"&monto&"',"
queryPagos=queryPagos&"'"&formapago&"',1,'"&rsUlt("UltFact")&"')"

conn.execute (queryPagos)
end if

'response.Write(query&" update AUTORIZACION set FACTURADO=0 WHERE ID_AUTORIZACION='"&idAuto&"' "&" "&queryPagos)
'response.End()

if err.number <> 0 then
	Response.Write("<commit>false</commit>")
else
	Response.Write("<commit>true</commit>")
end if
Response.Write("</insertar>") 
%>