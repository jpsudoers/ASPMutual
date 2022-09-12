<!--#include file="../../cnn_string.asp"-->
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

sql = "select ID_USUARIO as usuario,ID_INSTRUCTOR as relUser,dbo.MayMinTexto(NOMBRES+' '+A_PATERNO+' '+A_MATERNO) as nom_user from USUARIOS where ESTADO=1 and RUT='" + Request("RUT") + "' and CONTRASEÑA= '" + Request("PASS") + "' and ID_USUARIO not in (146)"

DATOS.Open sql,oConn

Response.Write("<?xml version='1.0' encoding='utf-8' ?>"&chr(13))
Response.Write("<rows>"&chr(13)) 
Response.Write("<records>"&DATOS.RecordCount&"</records>"&chr(13))

fila=0
WHILE NOT DATOS.EOF
	Response.Write("<row id="""&fila&""">"&chr(13))
	Response.Write("<USUARIO>"&DATOS("usuario")&"</USUARIO>"&chr(13))
	Response.Write("<NOM_USER>"&DATOS("nom_user")&"</NOM_USER>"&chr(13))
	Response.Write("</row>"&chr(13))

	fila=fila+1
	DATOS.MoveNext
WEND
Response.Write("</rows>") 
%>