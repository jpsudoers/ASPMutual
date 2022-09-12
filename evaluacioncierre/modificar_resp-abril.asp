<!--#include file="../conexion.asp"-->
<%
Response.CodePage = 65001
Response.CharSet = "utf-8"
Response.ContentType="text/xml"
Response.Write("<?xml version='1.0' encoding='UTF-8'?>")
Response.Write("<modificar>") 

on error resume next

For i = 1 To Request("countFilas") Step 1
query = "update HISTORICO_CURSOS set ASISTENCIA='"&Request("AsTra"&i)&"', "
query = query&"CALIFICACION='"&Request("caTra"&i)&"', EVALUACION='"&Request("eval"&i)&"', ESTADO=2 "
query = query&" where ID_HISTORICO_CURSO='"&Request("HiTra"&i)&"'"
conn.execute (query)
Next

Response.Write("<sql>"&query&"</sql>")
if err.number <> 0 then
	Response.Write("<commit>false</commit>")
else
	Response.Write("<commit>true</commit>")
end if
Response.Write("</modificar>") 
%>