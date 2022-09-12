<!--#include file="../../conexion.asp"-->
<%
Response.CodePage = 65001
Response.CharSet = "utf-8"
Response.ContentType="text/xml"
Response.Write("<?xml version='1.0' encoding='UTF-8'?>")
Response.Write("<insertar>") 

on error resume next
idAuto=request("idAuto")
usuario=request("usuario")
razon=request("razon")

query = "insert into VERIFICA_INSCRIPCION (ID_AUTORIZACION,RAZONES,FECHA_INGRESO,ID_USUARIO,ESTADO) values"
query = query&"('" & idAuto & "','" & razon & "',GETDATE(),'" & usuario & "',1)"

conn.execute (query)

conn.execute ("update AUTORIZACION set estado=1 WHERE ID_AUTORIZACION='"&idAuto&"'")

conn.execute ("update HISTORICO_CURSOS set estado=2 WHERE ID_AUTORIZACION='"&idAuto&"'")

'response.Write("update HISTORICO_CURSOS set estado=2 WHERE ID_AUTORIZACION='"&idAuto&"'")
'response.End()

if err.number <> 0 then
	Response.Write("<commit>false</commit>")
else
	Response.Write("<commit>true</commit>")
end if
Response.Write("</insertar>") 
%>