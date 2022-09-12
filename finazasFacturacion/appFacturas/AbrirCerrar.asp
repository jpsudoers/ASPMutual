<!--#include file="../../conexion.asp"-->
<%
Response.CodePage = 65001
Response.CharSet = "utf-8"
Response.ContentType="text/xml"
Response.Write("<?xml version='1.0' encoding='UTF-8'?>")
Response.Write("<insertar>") 

on error resume next
idusuario=request("usuario")

conn.execute ("UPDATE APERTURA_CIERRE SET FECHA_CIERRE=GETDATE(),ESTADO=0,USUARIO_CIERRE='"&idusuario&"' WHERE ESTADO=1")

conn.execute ("insert into APERTURA_CIERRE (FECHA_APERTURA,ESTADO,USUARIO) values(GETDATE(),1,'"&idusuario&"')")

'response.Write(query&" update AUTORIZACION set FACTURADO=0 WHERE ID_AUTORIZACION='"&idAuto&"' "&" "&queryPagos)
'response.End()

if err.number <> 0 then
	Response.Write("<commit>false</commit>")
else
	Response.Write("<commit>true</commit>")
end if
Response.Write("</insertar>") 
%>