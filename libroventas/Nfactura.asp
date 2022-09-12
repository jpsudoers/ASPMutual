<!--#include file="../conexion.asp"-->
<%
Response.CodePage = 65001
Response.CharSet = "utf-8"
Response.ContentType="text/xml"
Response.Write("<?xml version='1.0' encoding='UTF-8'?>")
Response.Write("<anular>") 

on error resume next

conn.execute ("update facturas set FACTURA="&Request("Nmfactura")&" where ID_FACTURA="&Request("idFactNota")&" and FACTURA is null")

conn.execute ("INSERT INTO FACTURAS_ASIGNACION_LOG (TIPO,ID_USUARIO,FECHA,ID_FACTURA) VALUES(1,"&Request("idUSRNota")&",GETDATE(),"&Request("idFactNota")&");")

if err.number <> 0 then
	Response.Write("<commit>false</commit>")
else
	Response.Write("<commit>true</commit>")
end if
Response.Write("</anular>") 
%>