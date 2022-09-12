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

sql = "select COUNT(PREINSCRIPCION_TRABAJADOR.id_trabajador) as total_trabajador from preinscripciones "
sql = sql&" inner join PREINSCRIPCION_TRABAJADOR on PREINSCRIPCION_TRABAJADOR.id_preinscripcion=preinscripciones.id_preinscripcion "
sql = sql&" where preinscripciones.id_preinscripcion="&Request("preinscripcion")

DATOS.Open sql,oConn

Response.Write("<?xml version='1.0' encoding='utf-8' ?>"&chr(13))
Response.Write("<totales>"&chr(13)) 

WHILE NOT DATOS.EOF
			Response.Write("<row>"&chr(13))
			Response.Write("<total>"&DATOS("total_trabajador")&"</total>"&chr(13))
			Response.Write("</row>"&chr(13))
	DATOS.MoveNext
WEND
Response.Write("</totales>") 
%>
