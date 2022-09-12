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

sql = "select TRABAJADOR.ID_TRABAJADOR,TRABAJADOR.RUT,TRABAJADOR.NOMBRES,TRABAJADOR.CARGO_EMPRESA,TRABAJADOR.ESCOLARIDAD,"
sql = sql&" TRABAJADOR.NOM_TRAB,TRABAJADOR.APATERTRAB,TRABAJADOR.AMATERTRAB,TRABAJADOR.CORREO,TRABAJADOR.EMAIL "
sql = sql&"  from TRABAJADOR where TRABAJADOR.RUT='"&Request("rut")&"'"

DATOS.Open sql,oConn

Response.Write("<?xml version='1.0' encoding='utf-8' ?>"&chr(13))
Response.Write("<DATOS>"&chr(13)) 

WHILE NOT DATOS.EOF
		Response.Write("<row>"&chr(13))
		Response.Write("<ID_TRABAJADOR>"&DATOS("ID_TRABAJADOR")&"</ID_TRABAJADOR>"&chr(13))
		Response.Write("<RUT>"&DATOS("RUT")&"</RUT>"&chr(13))
		Response.Write("<CARGO_EMPRESA>"&DATOS("CARGO_EMPRESA")&"</CARGO_EMPRESA>"&chr(13))
		Response.Write("<ESCOLARIDAD>"&DATOS("ESCOLARIDAD")&"</ESCOLARIDAD>"&chr(13))
		Response.Write("<NOM_TRAB>"&DATOS("NOM_TRAB")&"</NOM_TRAB>"&chr(13))
		Response.Write("<APATERTRAB>"&DATOS("APATERTRAB")&"</APATERTRAB>"&chr(13))
		Response.Write("<AMATERTRAB>"&DATOS("AMATERTRAB")&"</AMATERTRAB>"&chr(13))
		Response.Write("<CORREO>"&DATOS("CORREO")&"</CORREO>"&chr(13))
		Response.Write("<EMAIL>"&DATOS("EMAIL")&"</EMAIL>"&chr(13))
		Response.Write("</row>"&chr(13))
	DATOS.MoveNext
WEND
Response.Write("</DATOS>") 
%>
