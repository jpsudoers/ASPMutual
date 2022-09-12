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

vID = Request("id_autorizacion")

sql = "SELECT AUTORIZACION.ID_AUTORIZACION, EMPRESAS.ID_EMPRESA, EMPRESAS.R_SOCIAL, EMPRESAS.RUT, " 
sql = sql&"CURRICULO.NOMBRE_CURSO, CURRICULO.DESCRIPCION, AUTORIZACION.ID_PROGRAMA "
sql = sql&" FROM AUTORIZACION "
sql = sql&"inner join EMPRESAS on EMPRESAS.ID_EMPRESA=AUTORIZACION.ID_EMPRESA "
sql = sql&" inner join PROGRAMA on PROGRAMA.ID_PROGRAMA=AUTORIZACION.ID_PROGRAMA"
sql = sql&" inner join CURRICULO on CURRICULO.ID_MUTUAL=PROGRAMA.ID_MUTUAL"
sql = sql&" where AUTORIZACION.ID_AUTORIZACION='"&vID&"'"

DATOS.Open sql,oConn

Response.Write("<?xml version='1.0' encoding='utf-8' ?>"&chr(13))
Response.Write("<DATOS>"&chr(13)) 

WHILE NOT DATOS.EOF
		Response.Write("<row>"&chr(13))
		Response.Write("<IDEMPRESA>"&DATOS("ID_EMPRESA")&"</IDEMPRESA>"&chr(13))
		Response.Write("<RSOCIAL>"&DATOS("R_SOCIAL")&"</RSOCIAL>"&chr(13))
		Response.Write("<RUT>"&DATOS("RUT")&"</RUT>"&chr(13))
		Response.Write("<NOMBRE>"&DATOS("NOMBRE_CURSO")&"</NOMBRE>"&chr(13))
		Response.Write("<DESCRIPCION>"&DATOS("DESCRIPCION")&"</DESCRIPCION>"&chr(13))
		Response.Write("<IDMUTUAL>"&DATOS("ID_PROGRAMA")&"</IDMUTUAL>"&chr(13))
		Response.Write("</row>"&chr(13))
	DATOS.MoveNext
WEND
Response.Write("</DATOS>") 
%>
