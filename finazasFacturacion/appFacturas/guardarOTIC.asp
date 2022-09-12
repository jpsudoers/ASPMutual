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
formapago=request("formapago")
documento=request("documento")
ndocfact=request("ndocfact")
fechapago=request("fechapago")
montoEmp=request("montoEmp")
montoOtic=request("montoOtic")
txtNFotic=request("txtNFotic")
idotic=request("idotic")
fechaemision=request("fechaemision")
valCFact=request("valCFact") 
valCDesc=request("valCDesc")
desFact=request("desFact") 
valCurFact=request("valCurFact")

N_REG_SENCE="NULL"
if(request("N_REG_SENCE")<>"")then
	N_REG_SENCE="'"&request("N_REG_SENCE")&"'"
end if

if (cdbl(montoEmp) > 0)then
qEmp = "IF EXISTS (select * from AUTORIZACION where ID_AUTORIZACION='"&idAuto&"' and ESTADO=0 and FACTURADO=1) BEGIN "
qEmp=qEmp&"insert into FACTURAS (ID_AUTORIZACION,FECHA_EMISION,FECHA_VENCIMIENTO,ESTADO_CANCELACION,FACTURADO,ID_EMPRESA,"
qEmp=qEmp&"MONTO,TIPO_DOC,DOCUMENTO_COMPROMISO,N_DOCUMENTO,ESTADO,FECHA_FACTURA,VALOR_OC,"
qEmp=qEmp&"VALOR_C_DESC,DESCUENTO,VALOR_CURSO) values ('"&idAuto&"',"
qEmp=qEmp&"convert(date,'"&fechaemision&"'),CONVERT(datetime, CONVERT(varchar,convert(datetime,'"&fechaemision&"')+30,105),105),'"&pagada&"','"&usuario&"',"
qEmp=qEmp&"'"&idEmpresa&"','"&montoEmp&"','"&formapago&"','"&documento&"','"&ndocfact&"',1,GETDATE(),'" & valCFact &"','" & valCDesc &"','" & desFact &"','" & valCurFact &"') END"

conn.execute (qEmp)
end if

if (cdbl(montoOtic) > 0) then
qOTIC="IF EXISTS (select * from AUTORIZACION where ID_AUTORIZACION='"&idAuto&"' and ESTADO=0 and FACTURADO=1) BEGIN "
qOTIC=qOTIC&"insert into FACTURAS(ID_AUTORIZACION,FECHA_EMISION,FECHA_VENCIMIENTO,ESTADO_CANCELACION,FACTURADO,"
qOTIC=qOTIC&"ID_EMPRESA,MONTO,TIPO_DOC,DOCUMENTO_COMPROMISO,N_DOCUMENTO,ESTADO,FECHA_FACTURA,VALOR_OC,"
qOTIC=qOTIC&"VALOR_C_DESC,DESCUENTO,VALOR_CURSO) values ('"&idAuto&"',"
qOTIC=qOTIC&"convert(date,'"&fechaemision&"'),CONVERT(datetime, CONVERT(varchar,convert(datetime,'"&fechaemision&"')+30,105),105),'"&pagada&"','"&usuario&"',"
qOTIC=qOTIC&"'"&idotic&"','"&montoOtic&"','"&formapago&"','"&documento&"','"&ndocfact&"',1,GETDATE(),'" & valCFact &"','" & valCDesc &"','" & desFact &"','" & valCurFact &"') END"

conn.execute (qOTIC)
end if

'conn.execute ("update AUTORIZACION set FACTURADO=0 WHERE ID_AUTORIZACION='"&idAuto&"'")
conn.execute ("update AUTORIZACION set FACTURADO=0,N_REG_SENCE="&N_REG_SENCE&",ORDEN_COMPRA='"&ndocfact&"',VALOR_OC='"&valCDesc&"',VALOR_CURSO='"&valCurFact&"' WHERE ID_AUTORIZACION='"&idAuto&"'")

'response.Write(qOTIC&" "&qEmp)
'response.End()

if err.number <> 0 then
	Response.Write("<commit>false</commit>")
else
	Response.Write("<commit>true</commit>")
end if
Response.Write("</insertar>") 
%>