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
Dim DATOSOTIC
Dim oConn
SET oConn = Server.CreateObject("ADODB.Connection")
'oConn.Open("Provider=SQLOLEDB; User ID=sa;Password=SCL.2013.2013;data source=.\SQLEXPRESS;Initial Catalog=dbmas")
oConn.Open(MM_cnn_STRING)
Set DATOS = Server.CreateObject("ADODB.RecordSet")
DATOS.CursorType=3

sql = "select COUNT(*) AS 'exiFactura' from FACTURAS F where F.FACTURA='"&Request("n_fact")&"'"

DATOS.Open sql,oConn

Response.Write("<?xml version='1.0' encoding='utf-8' ?>"&chr(13))
Response.Write("<rows>"&chr(13)) 
Response.Write("<records>"&DATOS.RecordCount&"</records>"&chr(13))

fila=1
	Response.Write("<row id="""&fila&""">"&chr(13))
	Response.Write("<EXIFACT>"&DATOS("exiFactura")&"</EXIFACT>"&chr(13))
	Response.Write("</row>"&chr(13))
Response.Write("</rows>") 
%>