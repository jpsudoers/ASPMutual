<!--#include file="../conexion.asp"-->
<%
Response.CodePage = 65001
Response.CharSet = "utf-8"
Response.ContentType="text/xml"
Response.Write("<?xml version='1.0' encoding='UTF-8'?>")
Response.Write("<modificar>") 

on error resume next
vId  = Request("txtId")
vcurriculo = Request("Curriculo")
vinstructor = Request("Instructor")
vsede=Request("Sede")
vsence=Request("Sence")
vtipo=Request("Tipo")
vfechapertura=Request("txtFechApertura")
vfechcierre=Request("txtFechCierre")
vfechinicio=Request("txtFechInicio")
vfechtermino=Request("txtFechTermino")

vcupo=Request("txtCupo")
vinscritos=Request("txtInscritos")
vvacantes=Request("txtVacantes")

query = "UPDATE PROGRAMA SET ID_MUTUAL='"&vcurriculo&"', ID_INSTRUCTOR='"&vinstructor&"', ID_SEDE='"&vsede&"', SENCE='"&vsence&"'"
query = query&",TIPO='"&vtipo&"',FECHA_APERTURA=CONVERT(datetime,'"&vfechapertura&"',105),FECHA_CIERRE=CONVERT(datetime,'"&vfechcierre&"',105), "
query = query&"FECHA_INICIO_=CONVERT(datetime,'"&vfechinicio&"',105),FECHA_TERMINO=CONVERT(datetime,'"&vfechtermino&"',105), "
query = query&"CUPOS='"&vcupo&"', INSCRITOS='"&vinscritos&"', VACANTES='"&vvacantes&"' "
query = query&" WHERE ID_PROGRAMA = '"&vid&"'"

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