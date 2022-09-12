<%Response.CodePage = 65001
Response.CharSet = "utf-8"
Response.ContentType="text/xml"
Response.Write("<?xml version='1.0' encoding='UTF-8'?>")
Response.Write("<insertar>") %>
<!--#include file="../conexion.asp"-->
<%
on error resume next
vidempresa = Request("txtIdEmpresa")
vidprograma = Request("txtIdPrograma")
vdatostabla=Request("datostabla") 

cadena = vdatostabla
cadena = split(cadena, "/")

'for i = 0 to ubound(cadena)
dim query
query = "insert into HISTORICO_CURSOS (ID_EMPRESA,ID_PROGRAMA,ID_TRABAJADOR)"
query = query&" values('"&vidempresa&"','"&vidprograma&"','"&"28"&"') "
'RESPONSE.Write(query)
'RESPONSE.End()
'next

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