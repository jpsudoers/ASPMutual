<!--#include file="../conexion.asp"-->
<%
Response.CodePage = 65001
Response.CharSet = "utf-8"
Response.ContentType="text/xml"
Response.Write("<?xml version='1.0' encoding='UTF-8'?>")
Response.Write("<modificar>") 

on error resume next
vId = Request("OrdenC")
vfolio = Request("txFolio")
vfechemision = Request("txtFechEmision")
vfechvenc = Request("txtFechVenc")

query = "UPDATE AUTORIZACION SET FACTURA='"&vfolio&"', FECHA_EMISION=CONVERT(datetime,'"&vfechemision&"',105), "
query = query&"FECHA_VENCIMIENTO=CONVERT(datetime,'"&vfechvenc&"',105),ESTADO_CANCELACION=1 "
query = query&" WHERE ID_AUTORIZACION = '"&vId&"'"

'response.Write(query)
'response.End()

conn.execute (query)
Response.Write("<sql>"&query&"</sql>")
if err.number <> 0 then
	Response.Write("<commit>false</commit>")
else
	Response.Write("<commit>true</commit>")
end if
Response.Write("</modificar>") 
%>