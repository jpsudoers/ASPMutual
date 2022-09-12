<!--#include file="../conexion.asp"-->
<%Response.CodePage = 65001
Response.CharSet = "utf-8"
Response.ContentType="text/xml"
Response.Write("<?xml version='1.0' encoding='UTF-8'?>")
Response.Write("<eliminar>") 

on error resume next

dim vestado
if(Request("estado")="0")then
vestado="1"
else
vestado="0"
end if

dim query
query = "UPDATE EMPRESAS SET ACTIVA='"&vestado&"' WHERE ID_EMPRESA='"&Request("empresa")&"'"
Response.Write("<sql>"&query&"</sql>")
conn.execute (query)
if err.number <> 0 then
	Response.Write("<commit>false</commit>")
else
	Response.Write("<commit>true</commit>")
end if
Response.Write("</eliminar>") 
%>