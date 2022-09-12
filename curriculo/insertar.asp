<%Response.CodePage = 65001
Response.CharSet = "utf-8"
Response.ContentType="text/xml"
Response.Write("<?xml version='1.0' encoding='UTF-8'?>")
Response.Write("<insertar>") %>
<!--#include file="../conexion.asp"-->
<%
on error resume next
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

dim query
query = "IF NOT EXISTS (select * from CURRICULO C where C.CODIGO='"&vcod&"' and C.ESTADO=1) BEGIN "
query = query&"insert into CURRICULO (NOMBRE_CURSO, DESCRIPCION, OBJETIVOS, AUDIENCIA , HORAS, VIGENCIA, ESTADO, "
query = query&"CODIGO, SENCE, VALOR, VALOR_AFILIADOS, COD_SOFTLAND, ID_CECO) "
query = query&" values('"&vnomb&"','"&vdesc&"','"&vobj&"','"&vaud&"','"&vhor&"','"&vvig&"',1,'"&vcod&"'"
query = query&",'"&vsence&"','"&vvalor&"','"&vvalafil&"','"&vsoft&"','"&vceco&"') END"

'response.Write("<sql>"&query&"</sql>")
'response.End()

conn.execute (query)
if err.number <> 0 then
	Response.Write("<commit>false</commit>")
else
	Response.Write("<commit>true</commit>")
end if
Response.Write("<mensaje>"&vmensaje&"</mensaje>")
Response.Write("</insertar>") 
%>