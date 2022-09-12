<!--#include file="../conexion.asp"-->
<%
Response.CodePage = 65001
Response.CharSet = "utf-8"
Response.ContentType="text/xml"
Response.Write("<?xml version='1.0' encoding='UTF-8'?>")
Response.Write("<modificar>") 

on error resume next
vconantigua = Request("ConAntigua")
vnuevacont=Request("NuevaCont")


  query = "update EMPRESAS_USUARIOS set CONTRASENA='"&vnuevacont&"' where ID_EMPRESA='"&Session("usuario")
  query = query&"' and CONTRASENA='"&vconantigua&"' AND NOMBRE ='"&Session("correo_user")&"'"

conn.execute (query)
Response.Write("<sql>"&query&"</sql>")
if err.number <> 0 then
	Response.Write("<commit>false</commit>")
else
	Response.Write("<commit>true</commit>")
end if
Response.Write("</modificar>") 
%>