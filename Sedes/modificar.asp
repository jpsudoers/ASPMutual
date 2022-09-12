<!--#include file="../conexion.asp"-->
<%
Response.CodePage = 65001
Response.CharSet = "utf-8"
Response.ContentType="text/xml"
Response.Write("<?xml version='1.0' encoding='UTF-8'?>")
Response.Write("<modificar>") 

on error resume next
vId  = Request("txtId")
vnom=Request("txtNombre")
vdir=Request("txtDir")
vcom=Request("txtCom")
vciu=Request("txtCiu")

query = "UPDATE SEDES SET NOMBRE='"&vnom&"', DIRECCION='"&vdir&"', "
query = query&"COMUNA='"&vcom&"', CIUDAD='"&vciu&"' "
query = query&" WHERE ID_SEDE = '"&vid&"'"

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