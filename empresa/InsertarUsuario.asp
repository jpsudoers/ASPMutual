<!--#include file="../conexion.asp"-->
<%
Response.CodePage = 65001
Response.CharSet = "utf-8"
'Response.ContentType="text/xml"
Response.Write("<?xml version='1.0' encoding='UTF-8'?>")
Response.Write("<INSERTAR>") 


on error resume next

vid = Request("txtEmpresa")
vnomb=Request("txtNomb")
vmail=Request("txtMail") 
vcargo=Request("txtCargo")  
vfonocont=Request("txtFonoCont") 
vpasscord=Request("txtPassCord") 



query2 = "select top 1 RUT from EMPRESAS where ESTADO = 1 and ID_EMPRESA="&vid
set res = conn.execute(query2)

rutEmpresa = res("RUT")

vEstado = 1
query = "INSERT INTO EMPRESAS_USUARIOS (ID_EMPRESA, NOMBRE, EMAIL, TELEFONO, CARGO, CONTRASENA, ESTADO, RUT_EMPRESA)"
query = query&"values ("& vid & ",'" & vnomb & "','" & vmail & "','" & vfonocont & "','"& vcargo & "', '" &vpasscord & "'," & vEstado & ", "&rutEmpresa&")"


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