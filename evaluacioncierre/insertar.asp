<%Response.CodePage = 65001
Response.CharSet = "utf-8"
Response.ContentType="text/xml"
Response.Write("<?xml version='1.0' encoding='UTF-8'?>")
Response.Write("<insertar>") %>
<!--#include file="../conexion.asp"-->
<%
on error resume next
vcurriculo = Request("Curriculo")
vinstructor = Request("Instructor")
vsede=Request("Sede")
vsence=Request("Sence")
vtipo=Request("Tipo")
vfechapertura=Request("txtFechApertura")
vfechcierre=Request("txtFechCierre")
vfechinicio=Request("txtFechInicio")
vfechtermino=Request("txtFechTermino")
vvalor=Request("txtValor")
vcupo=Request("txtCupo")
vinscritos=Request("txtInscritos")
vvacantes=Request("txtVacantes")

dim query
query = "insert into PROGRAMA (ID_MUTUAL,ID_INSTRUCTOR,ID_SEDE,SENCE,TIPO,FECHA_APERTURA,FECHA_CIERRE"
query = query&",FECHA_INICIO_,FECHA_TERMINO,VALOR,CUPOS,INSCRITOS,VACANTES,ESTADO) "
query = query&" values('"&vcurriculo&"','"&vinstructor&"','"&vsede&"','"&vsence&"','"&vtipo&"',CONVERT(datetime,'"&vfechapertura&"',105)"
query = query&",CONVERT(datetime,'"&vfechcierre&"',105),CONVERT(datetime,'"&vfechinicio&"',105),CONVERT(datetime,'"&vfechtermino&"',105)"
query = query&",'"&vvalor&"','"&vcupo&"',0,'"&vvacantes&"',1) "

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