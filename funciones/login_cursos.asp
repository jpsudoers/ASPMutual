<!--#include file="../cnn_string.asp"-->
<%
Response.ContentType = "text/xml"
Response.AddHeader "Cache-control", "private"
Response.AddHeader "Expires", "-1"
Response.CodePage = 65001
Response.CharSet = "utf-8"
%>
<%
Dim DATOS
Dim oConn
SET oConn = Server.CreateObject("ADODB.Connection")
'oConn.Open("Provider=SQLOLEDB; User ID=sa;Password=SCL.2013.2013;data source=.\SQLEXPRESS;Initial Catalog=dbmas")
oConn.Open(MM_cnn_STRING)
Set DATOS = Server.CreateObject("ADODB.RecordSet")
DATOS.CursorType=3

vRutUsuario = Request("rut_usuario")
vPassUsuario = Request("pass")

sql = "select ID_USUARIO as usuario, NOMBRES+' '+A_PATERNO as nombre from USUARIOS "
sql = sql&" where RUT='"&vRutUsuario&"' and PERMISO5=1 and CONTRASEÑA='"&vPassUsuario&"'"

DATOS.Open sql,oConn

Response.Write("<?xml version='1.0' encoding='utf-8' ?>"&chr(13))
Response.Write("<DATOS>"&chr(13)) 

     if(DATOS.RecordCount<>"0")then
	    Session.TimeOut = 20
		Session("usuarioMutual") = DATOS("usuario")
		Session("nombre") = DATOS("nombre")

		Response.Write("<row>"&chr(13))
		Response.Write("<REGISTRO>"&DATOS.RecordCount&"</REGISTRO>"&chr(13))
		Response.Write("</row>"&chr(13))
	else
		Response.Write("<row>"&chr(13))
		Response.Write("<REGISTRO>"&DATOS.RecordCount&"</REGISTRO>"&chr(13))
		Response.Write("</row>"&chr(13))
	end if
Response.Write("</DATOS>") 
%>
