<!--#include file="../conexion.asp"-->
<%
Response.CodePage = 65001
Response.CharSet = "utf-8"
'Response.ContentType="text/xml"
Response.Write("<?xml version='1.0' encoding='UTF-8'?>")
Response.Write("<modificar>") 

on error resume next


vEmpre = Request("txtempre")
vUsuar = Request("txtusuario")
vid = Request("id")
vnomb=Request("txtNomb")
vmail=Request("txtMail") 
vcargo=Request("txtCargo")  

vfonocont=Request("txtFonoCont") 

vpasscord=Request("txtPassCord") 

vpassconta="NULL"
if(Request("txtPassConta")<>"")then
	vpassconta="LOWER('" & Request("txtPassConta") &"')"
end if

query = "update EMPRESAS_USUARIOS set NOMBRE = '"&vnomb&"', EMAIL ='"&vmail&"', TELEFONO = '"&vfonocont&"', CARGO = '"&vcargo&"', CONTRASENA = '"&vpasscord&"'"
query = query&" where ID_EMPRESAS_USUARIOS="&vUsuar


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