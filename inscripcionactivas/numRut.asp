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

DATOS.Open "SELECT RUT_TRAB FROM ASIGNA_RUT",oConn

Dim UPDATEDATOS
Set UPDATEDATOS = Server.CreateObject("ADODB.RecordSet")
UPDATEDATOS.CursorType=3

sql="UPDATE ASIGNA_RUT SET RUT_TRAB='"&CDBL(CDBL(DATOS("RUT_TRAB"))+CDBL(1))&"' WHERE RUT_TRAB='"&DATOS("RUT_TRAB")&"'"

UPDATEDATOS.Open sql,oConn

Response.Write("<?xml version='1.0' encoding='utf-8' ?>"&chr(13))
Response.Write("<DATOS>"&chr(13)) 
Response.Write("<row>"&chr(13))
Response.Write("<num>"&DATOS("RUT_TRAB")&"</num>"&chr(13))
Response.Write("</row>"&chr(13))
Response.Write("</DATOS>") 
%>
