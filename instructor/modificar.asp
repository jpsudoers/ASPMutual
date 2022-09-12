<!--#include file="../conexion.asp"-->
<%
Response.CodePage = 65001
Response.CharSet = "utf-8"
Response.ContentType="text/xml"
Response.Write("<?xml version='1.0' encoding='UTF-8'?>")
Response.Write("<modificar>") 

on error resume next
vId  = Request("txtId")
vrut = Request("txRut")
vnomb=Request("txtNomb")
vap_pater=Request("txtAp_pater")
vap_mater=Request("txtAp_mater")

query = "UPDATE INSTRUCTOR_RELATOR SET RUT='"&vrut&"', A_PATERNO='"&vap_pater&"', A_MATERNO='"&vap_mater&"', NOMBRES='"&vnomb&"' "
query = query&" WHERE ID_INSTRUCTOR = '"&vid&"'"

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