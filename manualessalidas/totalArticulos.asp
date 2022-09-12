<!--#include file="../cnn_string.asp"-->
<!--#include file="../conexion.asp"-->
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

sql = "SELECT ((select isnull(SUM(M1.CANTIDAD),0) from MOVIMIENTOS M1 "
sql = sql&" where M1.ID_ARTICULO='"&vID&"' and M1.TIPO_MOVIMIENTO in (1,2) and M1.MODULO=1) + " 
sql = sql&"(select isnull(SUM(M2.CANTIDAD),0) from MOVIMIENTOS M2 "
sql = sql&" where M2.ID_ARTICULO='"&vID&"' and M2.MODULO=3 and M2.TIPO_AJUSTE=1))- "
sql = sql&"((select isnull(SUM(M3.CANTIDAD),0) from MOVIMIENTOS M3 "
sql = sql&" where M3.ID_ARTICULO='"&vID&"' and M3.TIPO_MOVIMIENTO=3 and M3.MODULO=2)+ "
sql = sql&"(select isnull(SUM(M4.CANTIDAD),0) from MOVIMIENTOS M4 "
sql = sql&" where M4.ID_ARTICULO='"&vID&"' and M4.MODULO=3 and M4.TIPO_AJUSTE=2)) AS STOCK_ACTUAL_ARTICULO "

DATOS.Open sql,oConn

Response.Write("<?xml version='1.0' encoding='utf-8' ?>"&chr(13))
Response.Write("<DATOS>"&chr(13)) 

WHILE NOT DATOS.EOF
		Response.Write("<row>"&chr(13))
		Response.Write("<STOCK_ACTUAL_ARTICULO>"&DATOS("STOCK_ACTUAL_ARTICULO")&"</STOCK_ACTUAL_ARTICULO>"&chr(13))
		Response.Write("</row>"&chr(13))
	DATOS.MoveNext
WEND
Response.Write("</DATOS>") 
%>
