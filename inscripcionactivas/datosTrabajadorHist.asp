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

sql = "select ((select COUNT(*) from HISTORICO_CURSOS "
sql = sql&" where HISTORICO_CURSOS.ID_AUTORIZACION='"&Request("idauto")&"'"
sql = sql&" and HISTORICO_CURSOS.ID_TRABAJADOR='"&Request("trab")&"'"
sql = sql&" and HISTORICO_CURSOS.TRABUP<>'"&Request("trabfech")&"') + "
sql = sql&" (select COUNT(*) from HISTORICO_CURSOS  "
sql = sql&" where HISTORICO_CURSOS.ID_TRABAJADOR='"&Request("trab")&"'"
sql = sql&" and HISTORICO_CURSOS.ID_PROGRAMA='"&Request("id")&"') + "
sql = sql&" (select COUNT(*) from HISTORICO_CURSOS  "
sql = sql&" where HISTORICO_CURSOS.ID_AUTORIZACION='"&Request("idauto")&"'"
sql = sql&" and HISTORICO_CURSOS.TRABIDUP='"&Request("trab")&"'"
sql = sql&" and HISTORICO_CURSOS.TRABUP='"&Request("trabfech")&"')) as THisTrab "

DATOS.Open sql,oConn

Response.Write("<?xml version='1.0' encoding='utf-8' ?>"&chr(13))
Response.Write("<DATOS>"&chr(13)) 

WHILE NOT DATOS.EOF
		Response.Write("<row>"&chr(13))
		Response.Write("<THisTrab>"&DATOS("THisTrab")&"</THisTrab>"&chr(13))
		Response.Write("</row>"&chr(13))
	DATOS.MoveNext
WEND
Response.Write("</DATOS>") 
%>
