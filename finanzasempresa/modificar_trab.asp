<!--#include file="../conexion.asp"-->
<%
Response.CodePage = 65001
Response.CharSet = "utf-8"
Response.ContentType="text/xml"
Response.Write("<?xml version='1.0' encoding='UTF-8'?>")
Response.Write("<modificar>") 

'on error resume next

vNombre = trim(Request("txtNombre"))
vEmail = trim(Request("txtEmail"))
vFono = trim(Request("txtFono"))
vCargo = trim(Request("txtCargo"))
vComent = trim(Request("txtComent"))

dim query
if(Request("val1")="0") then
query = "UPDATE EMPRESAS SET NOMBRE_CONTA=dbo.MayMinTexto('"&vNombre&"'), EMAIL_CONTA=lower('"&vEmail&"'), "
query = query&"FONO_CONTABILIDAD='"&vFono&"', CARGO_CONTA=dbo.MayMinTexto('"&vCargo&"') WHERE ID_EMPRESA="&Request("ide")
else
query = "BEGIN UPDATE EMPRESA_USUARIOS SET NOMBRES=dbo.MayMinTexto('"&vNombre&"'), EMAIL=lower('"&vEmail&"'), "
query = query&"FONO='"&vFono&"', CARGO=dbo.MayMinTexto('"&vCargo&"') WHERE ID_EMP_USU="&Request("idc")&"; "
query = query&"UPDATE dbo.EMP_INS_USR SET COMENTARIO='"&vComent&"' WHERE ID_EMP_USU="&Request("idc")&" AND "
query = query&"ID_EMPRESA="&Request("ide")&"; END"
end if

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