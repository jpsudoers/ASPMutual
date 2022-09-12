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

vID = Request("id")

sql = "SELECT * FROM EMPRESAS "
sql = sql&" where EMPRESAS.ID_EMPRESA='"&vID&"'"

DATOS.Open sql,oConn

Response.Write("<?xml version='1.0' encoding='utf-8' ?>"&chr(13))
Response.Write("<DATOS>"&chr(13)) 

WHILE NOT DATOS.EOF
		Response.Write("<row>"&chr(13))
		Response.Write("<IDEMPRESA>"&DATOS("ID_EMPRESA")&"</IDEMPRESA>"&chr(13))
		Response.Write("<RSOCIAL>"&DATOS("R_SOCIAL")&"</RSOCIAL>"&chr(13))
		Response.Write("<DIRECCION>"&DATOS("DIRECCION")&"</DIRECCION>"&chr(13))
		Response.Write("<COMUNA>"&DATOS("COMUNA")&"</COMUNA>"&chr(13))
		Response.Write("<GIRO>"&DATOS("GIRO")&"</GIRO>"&chr(13))
		Response.Write("<FONO>"&DATOS("FONO")&"</FONO>"&chr(13))
		Response.Write("</row>"&chr(13))
	DATOS.MoveNext
WEND
Response.Write("</DATOS>") 
%>
