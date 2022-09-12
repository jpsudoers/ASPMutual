<%Response.CodePage = 65001
Response.CharSet = "utf-8"
Response.ContentType="text/xml"
Response.Write("<?xml version='1.0' encoding='UTF-8'?>")
Response.Write("<insertar>") %>
<!--#include file="../conexion.asp"-->
<%
on error resume next
vubi = Request("UbiBdg")
vdir = Request("DirBdg")
vres = Request("Responsable")

dim query
query = "insert into BODEGAS (UBICACION,DIRECCION,RESPONSABLE,ESTADO_BODEGA)"
query = query&" values('"&vubi&"','"&vdir&"','"&vres&"',1) "

conn.execute (query)

if err.number <> 0 then
	Response.Write("<commit>false</commit>")
else
	Response.Write("<commit>true</commit>")
end if
Response.Write("</insertar>") 
%>