<!--#include file="../conexion.asp"-->
<%
Response.CodePage = 65001
Response.CharSet = "utf-8"
Response.ContentType="text/xml"
Response.Write("<?xml version='1.0' encoding='UTF-8'?>")
Response.Write("<modificar>") 

on error resume next
vId  = Request("txtId")
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

query = "UPDATE TRABAJADOR SET RUT='"&vrut&"', A_PATERNO='"&vap_pater&"', A_MATERNO='"&vap_mater&"', NOMBRES='"&vnomb&"', "
query = query&"FECHA_NACIMIENTO=CONVERT(datetime,'"&vfnac&"',105), ESCOLARIDAD='"&vesc&"', ESPECIALIDAD='"&vesp&"', "
query = query&"DIRECCION='"&vdir&"', COMUNA='"&vcom&"', CIUDAD='"&vciu&"', FONO_FIJO='"&vfon&"', "
query = query&"CELULAR='"&vcel&"', EMAIL='"&vmail&"', CARGO_EMPRESA='"&vcargo&"'" 
query = query&" WHERE ID_TRABAJADOR = '"&vid&"'"

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