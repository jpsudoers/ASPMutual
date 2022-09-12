<!--#include file="../conexion.asp"-->
<%Response.CodePage = 65001
Response.CharSet = "utf-8"
Response.ContentType="text/xml"
Response.Write("<?xml version='1.0' encoding='UTF-8'?>")
Response.Write("<eliminar>") 

on error resume next
vid = Request("id") 

dim query
query = "UPDATE EMPRESAS_USUARIOS SET ESTADO=0 WHERE ID_EMPRESAS_USUARIOS="&vid
Response.Write("<sql>"&query&"</sql>")
conn.execute (query)

conn.execute ("insert into Log_Update_Insert values("&Session("usuarioMutual")&",getdate(),'"&Replace(query,"'","`")&"','"&Request.ServerVariables("PATH_INFO")&"','Usuarios de empresa');")

if err.number <> 0 then
	Response.Write("<commit>false</commit>")
else
	Response.Write("<commit>true</commit>")
end if
Response.Write("</eliminar>") 
%>