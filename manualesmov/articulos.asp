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

sql = "select A.ID_ARTICULO,A.DESC_ARTICULO from ARTICULOS A "
sql = sql&" inner join ARTICULO_BODEGA AR ON AR.ID_ARTICULO=A.ID_ARTICULO "
sql = sql&" WHERE A.ESTADO_ARTICULO=1 and AR.ESTADO_ARTICULO_BODEGA=1 AND AR.ID_BODEGA='"&Request("id_bodega")&"'"
sql = sql&" ORDER BY A.DESC_ARTICULO ASC "

DATOS.Open sql,oConn

Response.Write("<?xml version='1.0' encoding='utf-8' ?>"&chr(13))
Response.Write("<curriculo>"&chr(13)) 

WHILE NOT DATOS.EOF
		Response.Write("<row>"&chr(13))
		Response.Write("<ID_ARTICULO>"&DATOS("ID_ARTICULO")&"</ID_ARTICULO>"&chr(13))
		Response.Write("<DESC_ARTICULO>"&DATOS("DESC_ARTICULO")&"</DESC_ARTICULO>"&chr(13))
		Response.Write("</row>"&chr(13))
	DATOS.MoveNext
WEND
Response.Write("</curriculo>") 
%>
