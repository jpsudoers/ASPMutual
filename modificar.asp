<!--#include file="conexion.asp"-->
<%
Response.CodePage = 65001
Response.CharSet = "utf-8"
Response.ContentType="text/xml"
Response.Write("<?xml version='1.0' encoding='UTF-8'?>")
Response.Write("<modificar>") 

on error resume next
vId  = Request("txtId")
vrut = Request("txRut")
vrSocial=Request("txtRsoc")
vgiro=Request("txtGiro")
vdir=Request("txtDir")
vcom=Request("txtCom")
vciu=Request("txtCiu")
vfon=Request("txtFon")
vfax=Request("txtFax")
vmut=Request("txtMut")

vnomb=Request("txtNomb")
vmail=Request("txtMail") 
vcargo=Request("txtCargo")  

query = "UPDATE EMPRESAS SET RUT='"&vrut&"', R_SOCIAL='"&vrSocial&"', GIRO='"&vgiro&"', DIRECCION='"&vdir&"', "
query = query&"COMUNA='"&vcom&"', CIUDAD='"&vciu&"', FONO='"&vfon&"', FAX='"&vfax&"', "
query = query&"MUTUAL='"&vmut&"',NOMBRES='"&vnomb&"',EMAIL='"&vmail&"', CARGO='"&vcargo&"'" 
query = query&" WHERE id_empresa = '"&vid&"'"

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