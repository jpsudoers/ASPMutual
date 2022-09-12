<%Response.CodePage = 65001
Response.CharSet = "utf-8"
Response.ContentType="text/xml"
Response.Write("<?xml version='1.0' encoding='UTF-8'?>")
Response.Write("<insertar>") %>
<!--#include file="../conexion.asp"-->
<%
on error resume next
vnom=Request("txtNombre")
vdir=Request("txtDir")
vcom=Request("txtCom")
vciu=Request("txtCiu")

dim query
query = "insert into SEDES (ID_INSTITUCION,NOMBRE,DIRECCION,COMUNA,CIUDAD,ESTADO) "
query = query&" values(1,'"&vnom&"','"&vdir&"','"&vcom&"','"&vciu&"',1) "

Response.Write("<sql>"&query&"</sql>")
conn.execute (query)
if err.number <> 0 then
	Response.Write("<commit>false</commit>")
else
	Response.Write("<commit>true</commit>")
end if
Response.Write("<mensaje>"&vmensaje&"</mensaje>")
Response.Write("</insertar>") 
%>