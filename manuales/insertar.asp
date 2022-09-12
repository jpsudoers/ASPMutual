<%Response.CodePage = 65001
Response.CharSet = "utf-8"
Response.ContentType="text/xml"
Response.Write("<?xml version='1.0' encoding='UTF-8'?>")
Response.Write("<insertar>") %>
<!--#include file="../conexion.asp"-->
<%
on error resume next
vcodadic = Request("CodAdicArt")
vdesc = Request("DescArt")
vtipo = Request("TipoArt")
vumart = Request("UMArt")
vidprov = Request("IdProvArt")

dim query
query = "insert into ARTICULOS (CODIGO_ADICIONAL,DESC_ARTICULO,ESTADO_ARTICULO,ID_DPTO,UNIDAD_MEDIDA)"
query = query&" values('"&vcodadic&"','"&vdesc&"',1,'"&vtipo&"','"&vumart&"')"

conn.execute (query)

set rs = conn.execute ("select IDENT_CURRENT('ARTICULOS')AS UltART")

dim queryUp
queryUp = "UPDATE ARTICULO_BODEGA SET ID_ARTICULO='"&rs("UltART")&"', ID_ARTICULO_PROV=NULL "
queryUp = queryUp&" WHERE ID_ARTICULO_PROV='"&vidprov&"'"

conn.execute (queryUp)

if err.number <> 0 then
	Response.Write("<commit>false</commit>")
else
	Response.Write("<commit>true</commit>")
end if
Response.Write("</insertar>") 
%>