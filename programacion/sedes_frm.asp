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

sql = "select SEDES.ID_SEDE,dbo.MayMinTexto (SEDES.NOMBRE+', '+SEDES.DIRECCION+', '+SEDES.CIUDAD) as NOMBRE from SEDES "
sql = sql&" where SEDES.ID_SEDE not in(select bq.id_sede from bloque_programacion bq " 
sql = sql&" where bq.id_sede<>'"&Request("idSede")&"' and bq.id_programa='"&Request("progId")&"') "
sql = sql&" and SEDES.estado=1 order by SEDES.ID_SEDE asc"

DATOS.Open sql,oConn

Response.Write("<?xml version='1.0' encoding='utf-8' ?>"&chr(13))
Response.Write("<sede>"&chr(13)) 

WHILE NOT DATOS.EOF
		Response.Write("<row>"&chr(13))
		Response.Write("<ID>"&DATOS("ID_SEDE")&"</ID>"&chr(13))
		Response.Write("<NOMBRE>"&DATOS("NOMBRE")&"</NOMBRE>"&chr(13))
		Response.Write("</row>"&chr(13))
	DATOS.MoveNext
WEND
		Response.Write("<row>"&chr(13))
		Response.Write("<ID>27</ID>"&chr(13))
		Response.Write("<NOMBRE>Otras</NOMBRE>"&chr(13))
		Response.Write("</row>"&chr(13))
Response.Write("</sede>") 
%>
