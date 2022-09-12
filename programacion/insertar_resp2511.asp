<%Response.CodePage = 65001
Response.CharSet = "utf-8"
Response.ContentType="text/xml"
Response.Write("<?xml version='1.0' encoding='UTF-8'?>")
Response.Write("<insertar>") %>
<!--#include file="../conexion.asp"-->
<%
on error resume next
vcurriculo = Request("Curriculo")
vsence=Request("Sence")
vtipo=Request("Tipo")
vfechapertura=Request("txtFechApertura")
vfechcierre=Request("txtFechCierre")
vfechinicio=Request("txtFechInicio")
vfechtermino=Request("txtFechTermino")
vcupo=Request("txtCupo")
vinscritos=Request("txtInscritos")
vvacantes=Request("txtVacantes")
vprogId=Request("tabFecha")

dim query
query = "insert into PROGRAMA (ID_MUTUAL,SENCE,TIPO,FECHA_APERTURA,FECHA_CIERRE"
query = query&",FECHA_INICIO_,FECHA_TERMINO,CUPOS,INSCRITOS,VACANTES,ESTADO,VIGENCIA) "
query = query&" values('"&vcurriculo&"','"&vsence&"','"&vtipo&"',CONVERT(datetime,'"&vfechapertura&"',105)"
query = query&",CONVERT(datetime,'"&vfechcierre&"',105),CONVERT(datetime,'"&vfechinicio&"',105),CONVERT(datetime,'"&vfechtermino&"',105)"
query = query&",'"&vcupo&"',0,'"&vvacantes&"',1,1) "

conn.execute (query)

set rsProg = conn.execute ("select IDENT_CURRENT('PROGRAMA')AS UltProg")

conn.execute ("update bloque_programacion set id_programa='"&rsProg("UltProg")&"' where id_programa='"&vprogId&"'")

Response.Write("<sql>"&query&"</sql>")

if err.number <> 0 then
	Response.Write("<commit>false</commit>")
else
	Response.Write("<commit>true</commit>")
end if
Response.Write("<mensaje>"&vmensaje&"</mensaje>")
Response.Write("</insertar>") 
%>