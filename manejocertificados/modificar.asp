<!--#include file="../conexion.asp"-->
<%
Response.CodePage = 65001
Response.CharSet = "utf-8"
Response.ContentType="text/xml"
Response.Write("<?xml version='1.0' encoding='UTF-8'?>")
Response.Write("<modificar>") 

on error resume next

For i = 0 To Request("countFilas") Step 1
query = "update HISTORICO_CURSOS set ASISTENCIA='"&Request("Asistencia"&i)&"', "
query = query&"CALIFICACION='"&Request("calificacion"&i)&"', EVALUACION='"&Request("eval"&i)&"', ESTADO=0 "
query = query&" where ID_HISTORICO_CURSO='"&Request("Historico"&i)&"'"
conn.execute (query)
Next

conn.execute ("UPDATE AUTORIZACION SET ESTADO=0 WHERE AUTORIZACION.ID_PROGRAMA="&Request("txtId"))

conn.execute ("UPDATE PROGRAMA SET VIGENCIA=0 WHERE ID_PROGRAMA="&Request("txtId"))
'response.Write(query)
'response.End()

Response.Write("<sql>"&query&"</sql>")
if err.number <> 0 then
	Response.Write("<commit>false</commit>")
else
	Response.Write("<commit>true</commit>")
end if
Response.Write("</modificar>") 
%>