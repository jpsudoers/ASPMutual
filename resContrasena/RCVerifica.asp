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

vempresa = trim(Request("emp"))
vcorreo = trim(Request("crr"))

query = "select(select COUNT(*) from EMPRESAS where RUT='"&vempresa&"' and EMAIL='"&vcorreo&"' "
query = query&" and PASSWORD_COORDINACION is not null and PASSWORD_COORDINACION<>'' and PREINSCRITA=1) as cordinacion, "
query = query&"(select COUNT(*) from EMPRESAS  where RUT='"&vempresa&"' and EMAIL_CONTA='"&vcorreo&"' "
query = query&" and PASSWORD_CONTA is not Null  and PASSWORD_CONTA<>'' and PREINSCRITA=1) as contabilidad "

DATOS.Open query,oConn

Response.Write("<?xml version='1.0' encoding='utf-8' ?>"&chr(13))
Response.Write("<DATOS>"&chr(13)) 

WHILE NOT DATOS.EOF
		Response.Write("<row>"&chr(13))
		Response.Write("<VMRcoord>"&DATOS("cordinacion")&"</VMRcoord>"&chr(13))
		Response.Write("<VMRconta>"&DATOS("contabilidad")&"</VMRconta>"&chr(13))
		Response.Write("</row>"&chr(13))
	DATOS.MoveNext
WEND
Response.Write("</DATOS>") 
%>
