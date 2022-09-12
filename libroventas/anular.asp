<!--#include file="../conexion.asp"-->
<%
Response.CodePage = 65001
Response.CharSet = "utf-8"
Response.ContentType="text/xml"
Response.Write("<?xml version='1.0' encoding='UTF-8'?>")
Response.Write("<anular>") 

on error resume next

if(Request("tipo_baja")="0")then
	tipo="0"
else
	tipo="2"
end if

conn.execute ("update AUTORIZACION set FACTURADO=1 where ID_AUTORIZACION='"&Request("idAutoAnular")&"' and FACTURADO=0")

anulaFac="update FACTURAS set ESTADO='"&tipo&"', FECHA_ESTADO=GETDATE(), DESCRIPCION_ESTADO='"&Request("razonAnular")&"' "
anulaFac=anulaFac&"where ID_AUTORIZACION='"&Request("idAutoAnular")&"' and ID_FACTURA='"&Request("idFactAnular")&"'"

conn.execute (anulaFac)
'response.Write(anulaFac&" "&"update AUTORIZACION set FACTURADO=1 where ID_AUTORIZACION='"&Request("idAutoAnular")&"' and FACTURADO=0")

if err.number <> 0 then
	Response.Write("<commit>false</commit>")
else
	Response.Write("<commit>true</commit>")
end if
Response.Write("</anular>") 
%>