<!--#include file="../conexion.asp"-->
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
'oConn.Open("Provider=SQLOLEDB; User ID=sa;Password=SCL.2013.2013;data source=.\SQLEXPRESS;Initial Catalog=dbarauco")
oConn.Open(MM_cnn_STRING)
Set DATOS = Server.CreateObject("ADODB.RecordSet")
DATOS.CursorType=3

vID = Request("id")

sql = "select P.ID_PROVEEDORES,P.RUT,P.PROVEEDOR from PROVEEDORES p where p.ID_PROVEEDORES='"&vID&"'"

DATOS.Open sql,oConn

Response.Write("<?xml version='1.0' encoding='utf-8' ?>"&chr(13))
Response.Write("<DATOS>"&chr(13)) 

WHILE NOT DATOS.EOF
		Response.Write("<row>"&chr(13))
		Response.Write("<IDPROV>"&DATOS("ID_PROVEEDORES")&"</IDPROV>"&chr(13))
		Response.Write("<RUT>"&DATOS("RUT")&"</RUT>"&chr(13))
		Response.Write("<PROVEEDOR>"&DATOS("PROVEEDOR")&"</PROVEEDOR>"&chr(13))
		Response.Write("</row>"&chr(13))
	DATOS.MoveNext
WEND
Response.Write("</DATOS>") 
%>
