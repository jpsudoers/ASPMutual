<!--#include file="../conexion.asp"-->
<%
Response.CodePage = 65001
Response.CharSet = "utf-8"
Response.ContentType="text/xml"
Response.Write("<?xml version='1.0' encoding='UTF-8'?>")
Response.Write("<modificar>") 

on error resume next
vbloque=Request("txtBloque")
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

query = "UPDATE bloque_programacion SET id_relator='"&vrelator&"', id_sede='"&vsede&"',cupos='"&vcupo&"',"
query = query&"id_rel_seg="&vrelator_seg&",nom_sede="&vtxtdir&",n_relatores="&vn_relator
query = query&" WHERE id_programa='"&vprograma&"' and id_bloque='"&vbloque&"'"

conn.execute (query)

UpHis = "update HISTORICO_CURSOS set RELATOR='"&vrelator&"', SEDE='"&vsede&"'"
UpHis = UpHis&" where ID_BLOQUE='"&vbloque&"' and ID_PROGRAMA='"&vprograma&"'" 

conn.execute (UpHis)

Response.Write("<sql>"&query&"</sql>")
if err.number <> 0 then
	Response.Write("<commit>false</commit>")
else
	Response.Write("<commit>true</commit>")
end if
Response.Write("</modificar>") 
%>