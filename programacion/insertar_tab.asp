<%Response.CodePage = 65001
Response.CharSet = "utf-8"
Response.ContentType="text/xml"
Response.Write("<?xml version='1.0' encoding='UTF-8'?>")
Response.Write("<insertar>") %>
<!--#include file="../conexion.asp"-->
<%
on error resume next
vprograma=Request("tabPrograma")
vrelator=Request("relator_frm")
vsede=Request("salaSede_frm")
vcupo=Request("txtCupo_frm")

     
if(Request("relator_frm_seg")<>"")then
	vrelator_seg="'"&Request("relator_frm_seg")&"'"
else
	vrelator_seg="NULL"
end if

if(Request("txtDir_frm")<>"")then
	vtxtdir="dbo.MayMinTexto('"&trim(Request("txtDir_frm"))&"')"
else
	vtxtdir="NULL"
end if

vn_relator=Request("txtNRelator")

dim query
query="IF NOT EXISTS (select * from bloque_programacion bp where bp.id_programa='"&vprograma&"'"
query=query&" and bp.id_relator='"&vrelator&"') BEGIN "
query=query&"INSERT INTO bloque_programacion (id_programa,id_relator,id_sede,cupos,estado,id_rel_seg,nom_sede,n_relatores) "
query = query&" values('"&vprograma&"','"&vrelator&"','"&vsede&"','"&vcupo&"',1,"&vrelator_seg&","&vtxtdir&","&vn_relator&") END "

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