<!--#include file="../conexion.asp"-->
<%
Response.CodePage = 65001
Response.CharSet = "utf-8"
Response.ContentType="text/xml"
Response.Write("<?xml version='1.0' encoding='UTF-8'?>")
Response.Write("<modificar>") 

on error resume next
vruttrab = Request("txtRutTrab")
vcargotrab = trim(Request("txtCargoTrab"))
vescolaridadtrab = Request("escolaridadTrab")
vnomtrab = trim(Request("txtNomTrab"))
vapatertrab = trim(Request("txtAPaterTrab"))
vamatertrab = trim(Request("txtAMaterTrab"))
vtrabid = Request("txtTrabID")
VCORREO = Request("txtMail")

emailTrab = "Null"
if(Request("txtEmailTrab")<>"")then
	emailTrab = "'"&Request("txtEmailTrab")&"'"
end if

query = "UPDATE TRABAJADOR SET NOMBRES=dbo.MayMinTexto('"&vnomtrab&" "&vapatertrab&" "&vamatertrab&"'),"
query = query&"CARGO_EMPRESA=dbo.MayMinTexto('"&vcargotrab&"'), ESCOLARIDAD='"&vescolaridadtrab&"',"
query = query&"NOM_TRAB=dbo.MayMinTexto('"&vnomtrab&"'),APATERTRAB=dbo.MayMinTexto('"&vapatertrab&"'),"
query = query&"AMATERTRAB=dbo.MayMinTexto('"&vamatertrab&"'),CORREO='"&VCORREO&"',EMAIL="&emailTrab&" WHERE ID_TRABAJADOR = '"&vtrabid&"'"

conn.execute (query)
Response.Write("<sql>"&query&"</sql>")
if err.number <> 0 then
	Response.Write("<commit>false</commit>")
else
	Response.Write("<commit>true</commit>")
end if
Response.Write("</modificar>") 
%>