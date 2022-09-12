<%Response.CodePage = 65001
Response.CharSet = "utf-8"
Response.ContentType="text/xml"
Response.Write("<?xml version='1.0' encoding='UTF-8'?>")
Response.Write("<insertar>") %>
<!--#include file="../conexion.asp"-->
<%
on error resume next
vempresa = Request("Empresa")
vcurso=Request("txtCurso")
vfechrequerida=Request("txtFecha")
vparticipantes=Request("txtPart")

dim query2
query2 = "insert into SOLICITUD (id_empresa,id_mutual,fecha,estado,participantes"
query2 = query2&") values('"&vempresa&"','"&vcurso&"',CONVERT(datetime,'"&vfechrequerida&"',105), "
query2 = query2&"1,'"&vparticipantes&"') "

conn.execute (query2)

if err.number <> 0 then
	Response.Write("<commit>false</commit>")
else
	Response.Write("<commit>true</commit>")
end if
Response.Write("<mensaje>"&vmensaje&"</mensaje>")
Response.Write("</insertar>") 
%>