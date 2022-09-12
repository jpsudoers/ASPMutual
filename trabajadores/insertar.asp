<%Response.CodePage = 65001
Response.CharSet = "utf-8"
Response.ContentType="text/xml"
Response.Write("<?xml version='1.0' encoding='UTF-8'?>")
Response.Write("<insertar>") %>
<!--#include file="../conexion.asp"-->
<%
on error resume next
vrut = Request("txRut")
vnomb=Request("txtNomb")
vap_pater=Request("txtAp_pater")
vap_mater=Request("txtAp_mater")

vesc=Request("txtEscol")
vesp=Request("txtEsp")
vfnac=Request("txtFecha")

vdir=Request("txtDir")
vcom=Request("txtCom")
vciu=Request("txtCiu")
vfon=Request("txtFon")
vcel=Request("txtCel")
vmail=Request("txtMail") 
vcargo=Request("txtCargo")  

dim query
query = "insert into TRABAJADOR (RUT, A_PATERNO, A_MATERNO, NOMBRES, FECHA_NACIMIENTO, ESCOLARIDAD, ESPECIALIDAD, DIRECCION, COMUNA"
query = query&",CIUDAD, FONO_FIJO, CELULAR, EMAIL, CARGO_EMPRESA,ESTADO) "
query = query&" values('"&vrut&"','"&vap_pater&"','"&vap_mater&"','"&vnomb&"',CONVERT(datetime,'"&vfnac&"',105),'"&vesc&"','"&vesp&"','"&vdir&"','"&vcom&"',"
query = query&"'"&vciu&"','"&vfon&"','"&vcel&"','"&vmail&"','"&vcargo&"',1) "

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