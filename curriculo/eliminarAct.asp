<!--#include file="../conexion.asp"-->
<%Response.CodePage = 65001
Response.CharSet = "utf-8"
Response.ContentType="text/xml"
Response.Write("<?xml version='1.0' encoding='UTF-8'?>")
Response.Write("<eliminar>") 

on error resume next
vid = Request("id") 
vtxtBloqueId = Request("b") 

dim query
query = "update CURRICULO_ACTIVIDADES set ESTADO=0 where ID_ACTIVIDAD_CURSO='"&vid&"';"
query = query&"IF NOT EXISTS (select * from CURRICULO_ACTIVIDADES ac where ac.ID_BLOQUE_CURSO='"&vtxtBloqueId&"' and ac.estado=1) "
query = query&" BEGIN UPDATE CURRICULO_BLOQUE SET ESTADO=0 WHERE ID_BLOQUE_CURSO='"&vtxtBloqueId&"' END"

Response.Write("<sql>"&query&"</sql>")
conn.execute (query)
if err.number <> 0 then
	Response.Write("<commit>false</commit>")
else
	Response.Write("<commit>true</commit>")
end if
Response.Write("</eliminar>") 
%>