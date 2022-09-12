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



dim query2
query2 ="select ID_BANCO, ID_CUENTA_CORRIENTE, NUMERO_CUENTA "
query2 = query2&" from CUENTAS_CORRIENTES "
query2 = query2&" WHERE ID_BANCO ='"&Request("id_banco")&"'"


DATOS.Open query2,oConn

Response.Write("<?xml version='1.0' encoding='utf-8' ?>"&chr(13))
Response.Write("<instructor>"&chr(13)) 

WHILE NOT DATOS.EOF
		Response.Write("<row>"&chr(13))
		Response.Write("<ID_CUENTA_CORRIENTE>"&DATOS("ID_CUENTA_CORRIENTE")&"</ID_CUENTA_CORRIENTE>"&chr(13))
		Response.Write("<NUMERO_CUENTA>"&DATOS("NUMERO_CUENTA")&"</NUMERO_CUENTA>"&chr(13))
		Response.Write("</row>"&chr(13))
	DATOS.MoveNext
WEND
Response.Write("</instructor>") 
%>
