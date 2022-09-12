<!--#include file="../conexion.asp"-->
<%
Response.CodePage = 65001
Response.CharSet = "utf-8"
Response.ContentType="text/xml"
Response.Write("<?xml version='1.0' encoding='UTF-8'?>")
Response.Write("<modificar>") 

on error resume next
vId  = Request("txtIdIngresos")
vperido=Request("Perido")
vcontratista=Request("Contratista")
vcomp=Request("txComp")

query = "UPDATE ingresos SET "
query = query&"periodo='"&vperido&"', contabilizado='"&vcontratista&"',comprobante='"&vcomp&"' "
query = query&" WHERE id_ingresos = '"&vid&"'"

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