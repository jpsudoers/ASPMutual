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
query = "BEGIN INSERT INTO dbo.EMPRESA_USUARIOS (NOMBRES, EMAIL, FONO, CARGO) VALUES (dbo.MayMinTexto('"&vNombre&"'), "
query = query&"lower('"&vEmail&"'), '"&vFono&"', dbo.MayMinTexto('"&vCargo&"')); "
query = query&"INSERT INTO dbo.EMP_INS_USR (ID_EMP_USU, FECHA_CREACION, ESTADO, TIPO_CONTACTO, ID_EMPRESA, COMENTARIO) VALUES "
query = query&"((SELECT MAX(ID_EMP_USU) FROM dbo.EMPRESA_USUARIOS), GETDATE(), 1, 4, "&Request("ide")&", '"&vComent&"'); END"

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