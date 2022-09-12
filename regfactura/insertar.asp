<%Response.CodePage = 65001
Response.CharSet = "utf-8"
Response.ContentType="text/xml"
Response.Write("<?xml version='1.0' encoding='UTF-8'?>")
Response.Write("<insertar>") %>
<!--#include file="../conexion.asp"-->
<%
on error resume next
vidempresa = Request("txtIdEmpresa")
vfolio=Request("txFolio")
vfechemision=Request("txtFechEmision")
vfechvenc=Request("txtFechVenc")

dim query
query = "insert into FACTURAS (id_empresa,folio,fecha_emision,fecha_entrega"
query = query&",estado) values('"&vidempresa&"','"&vfolio&"',CONVERT(datetime,'"&vfechemision&"',105),CONVERT(datetime,'"&vfechvenc&"',105),1) "

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