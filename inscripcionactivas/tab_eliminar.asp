<%Response.CodePage = 65001
Response.CharSet = "utf-8"
Response.ContentType="text/xml"
Response.Write("<?xml version='1.0' encoding='UTF-8'?>")
Response.Write("<insertar>") %>
<!--#include file="../conexion.asp"-->
<%
on error resume next
vhistid = Request("histId")
vtabpart = Request("tabPart")

dim delHistTrab
delHistTrab = "update HISTORICO_CURSOS set TRABDEL='"&vtabpart&"'"
delHistTrab = delHistTrab&" where ID_HISTORICO_CURSO='"&vhistid&"'"

conn.execute (delHistTrab)

if err.number <> 0 then
	Response.Write("<commit>false</commit>")
else
	Response.Write("<commit>true</commit>")
end if
Response.Write("</insertar>") 
%>