<!--#include file="../conexion.asp"-->
<%
Response.CodePage = 65001
Response.CharSet = "utf-8"
Response.ContentType="text/xml"
Response.Write("<?xml version='1.0' encoding='UTF-8'?>")
Response.Write("<modificar>") 

on error resume next
vId = Request("IdBdg")
vubi = Request("UbiBdg")
vdir = Request("DirBdg")
vres = Request("Responsable")

query = "UPDATE BODEGAS SET UBICACION='"&vubi&"', DIRECCION='"&vdir&"', RESPONSABLE='"&vres&"'"
query = query&" WHERE ID_BODEGA='"&vid&"'"

conn.execute (query)

if err.number <> 0 then
	Response.Write("<commit>false</commit>")
else
	Response.Write("<commit>true</commit>")
end if
Response.Write("</modificar>") 
%>