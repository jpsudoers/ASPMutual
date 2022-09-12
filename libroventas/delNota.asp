<!--#include file="../conexion.asp"-->
<%
Response.CodePage = 65001
Response.CharSet = "utf-8"
Response.ContentType="text/xml"
Response.Write("<?xml version='1.0' encoding='UTF-8'?>")
Response.Write("<anular>") 

on error resume next

conn.execute ("exec [dbo].[Elimina_Nota_Venta] "&Request("idFactNota")&", "&Request("idAutoNota")&", "&Request("tipo_bajaNota")&", "&Request("idUSRNota")&", '"&Request("razonDelNota")&"'")



if err.number <> 0 then
	Response.Write("<commit>false</commit>")
else
	Response.Write("<commit>true</commit>")
end if
Response.Write("</anular>") 
%>