<%Response.CodePage = 65001
Response.CharSet = "utf-8"
Response.ContentType="text/xml"
Response.Write("<?xml version='1.0' encoding='UTF-8'?>")
Response.Write("<insertar>") %>
<!--#include file="conexion.asp"-->
<%
on error resume next
vrut = Request("rut")
vnombre = Request("nombreCompleto")
vcorreoElectronico = Request("correoElectronico")


dim query
query = " INSERT INTO prueba(nombre,rut,correoElectronico) VALUES ('"&vnombre&"','"&vrut&"','"&vcorreoElectronico&"');"

'response.Write(query)
'response.End()

'Response.Write("<sql>"&query&"</sql>")
conn.execute (query)
if err.number <> 0 then
	Response.Write("<commit>false</commit>")
else
	Response.Write("<commit>true</commit>")
end if
'Response.Write("<mensaje>"&vmensaje&"</mensaje>")
Response.Write("</insertar>") 
%>