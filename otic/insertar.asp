<%Response.CodePage = 65001
Response.CharSet = "utf-8"
Response.ContentType="text/xml"
Response.Write("<?xml version='1.0' encoding='UTF-8'?>")
Response.Write("<insertar>") %>
<!--#include file="../conexion.asp"-->
<%
on error resume next
vrut = Request("txRut")
vrSocial=Request("txtRsoc")
vgiro=Request("txtGiro")
vdir=Request("txtDir")
vcom=Request("txtCom")
vciu=Request("txtCiu")
vfon=Request("txtFon")
vfax=Request("txtFax")

vnomb=Request("txtNomb")
vmail=Request("txtMail") 
vcargo=Request("txtCargo")  

dim query
query = "insert into EMPRESAS (RUT,R_SOCIAL ,GIRO ,DIRECCION,COMUNA,CIUDAD ,FONO ,FAX , ESTADO "
query = query&",NOMBRES ,EMAIL "
query = query&",CARGO, TIPO) values('"&vrut&"','"&vrSocial&"','"&vgiro&"','"&vdir&"','"&vcom&"','"&vciu&"','"&vfon&"','"&vfax&"', "
query = query&"1,'"&vnomb&"','"&vmail&"','"&vcargo&"',2) "

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