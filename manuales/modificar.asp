<!--#include file="../conexion.asp"-->
<%
Response.CodePage = 65001
Response.CharSet = "utf-8"
Response.ContentType="text/xml"
Response.Write("<?xml version='1.0' encoding='UTF-8'?>")
Response.Write("<modificar>") 

on error resume next
vId = Request("IdProvArt")
vcodadic = Request("CodAdicArt")
vdesc = Request("DescArt")
vtipo = Request("TipoArt")
vumart = Request("UMArt")

query = "UPDATE ARTICULOS SET CODIGO_ADICIONAL='"&vcodadic&"', DESC_ARTICULO='"&vdesc&"',"
query = query&" ID_DPTO='"&vtipo&"', UNIDAD_MEDIDA='"&vumart&"' "
query = query&" WHERE ID_ARTICULO='"&vid&"'"

conn.execute (query)

dim queryUp
queryUp = "UPDATE ARTICULO_BODEGA SET ESTADO_ARTICULO_BODEGA=0, ESTADO_REGISTRO=NULL "
queryUp = queryUp&" WHERE ID_ARTICULO='"&vId&"' and ESTADO_REGISTRO='"&Request("EstBdg")&"'"

conn.execute (queryUp)

if err.number <> 0 then
	Response.Write("<commit>false</commit>")
else
	Response.Write("<commit>true</commit>")
end if
Response.Write("</modificar>") 
%>