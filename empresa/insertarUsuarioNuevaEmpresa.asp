<!--#include file="../conexion.asp"-->
<%
Response.CodePage = 65001
Response.CharSet = "utf-8"
'Response.ContentType="text/xml"



vid = Request("txtrut")
vnomb=Request("txtNomb")
vmail=Request("txtMail") 
vcargo=Request("txtCargo")  
vfonocont=Request("txtFonoCont") 
vpasscord=Request("txtPassCord") 
vidEmpre = 0
vEstado = 1

query = "INSERT INTO EMPRESAS_USUARIOS (ID_EMPRESA, NOMBRE, EMAIL, TELEFONO, CARGO, CONTRASENA, ESTADO, RUT_EMPRESA)"
query = query&"values ("& vidEmpre & ",'" & vnomb & "','" & vmail & "','" & vfonocont & "','"& vcargo & "', '" &vpasscord & "'," & vEstado & ", '"&vid&"')"




conn.execute (query)

Response.Write("<sql>"&query&"</sql>")
if err.number <> 0 then
	Response.Write("<commit>false</commit>")
else
	Response.Write("<commit>true</commit>")
end if
Response.Write("</modificar>") 






%>