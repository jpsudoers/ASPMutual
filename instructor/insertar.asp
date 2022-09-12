<%Response.CodePage = 65001
Response.CharSet = "utf-8"
Response.ContentType="text/xml"
Response.Write("<?xml version='1.0' encoding='UTF-8'?>")
Response.Write("<insertar>") %>
<!--#include file="../conexion.asp"-->
<%
on error resume next
vrut = Request("txRut")
vnomb=Request("txtNomb")
vap_pater=Request("txtAp_pater")
vap_mater=Request("txtAp_mater")

dim query
query = "IF NOT EXISTS (select * from INSTRUCTOR_RELATOR i where i.rut='"&vrut&"' and i.ESTADO=1) BEGIN "
query = query&"insert into INSTRUCTOR_RELATOR (RUT, A_PATERNO, A_MATERNO, NOMBRES, ESTADO) "
query = query&" values('"&vrut&"','"&vap_pater&"','"&vap_mater&"','"&vnomb&"',1) END"

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