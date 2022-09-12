<!--#include file="../conexion.asp"-->
<%
Response.CodePage = 65001
Response.CharSet = "utf-8"
Response.ContentType="text/xml"
Response.Write("<?xml version='1.0' encoding='UTF-8'?>")
Response.Write("<modificar>") 

on error resume next
vId  = Request("txtId")
vnomb=Request("txtNom")
vdesc = Request("txtDesc")
vobj=Request("txtObj")
vaud=Request("txtAud")
vhor=Request("txtHor")
vvig=Request("txtVig") 
vcod=Request("txtCod")
vsence=Request("Sence")
vvalor=Request("txtValor")
vvalafil=Request("txtValAfiliados")
vsoft=Request("txtSoftland")
vceco=Request("txtCeco")
vAsistencia=Request("txtAsistencia")
vCalificacion=Request("txtCalificacion")


query = "UPDATE CURRICULO SET NOMBRE_CURSO='"&vnomb&"', DESCRIPCION='"&vdesc&"', OBJETIVOS='"&vobj&"', AUDIENCIA='"&vaud&"', "
query = query&"HORAS='"&vhor&"', VIGENCIA='"&vvig&"', CODIGO='"&vcod&"', SENCE='"&vsence&"', VALOR='"&vvalor&"', "
query = query&" VALOR_AFILIADOS='"&vvalafil&"', COD_SOFTLAND='"&vsoft&"', ID_CECO='"&vceco&"', PORCE_ASISTENCIA= '"&vAsistencia&"', PORCE_CALIFICACION='"&vCalificacion&"' WHERE ID_MUTUAL = '"&vid&"'"

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