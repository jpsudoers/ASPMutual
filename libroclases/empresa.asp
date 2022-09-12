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

vId=Request("id_programa")
sql = "select distinct EMPRESAS.ID_EMPRESA,EMPRESAS.R_SOCIAL from HISTORICO_CURSOS "
sql = sql&" inner join EMPRESAS on EMPRESAS.ID_EMPRESA=HISTORICO_CURSOS.ID_EMPRESA "
sql = sql&" where HISTORICO_CURSOS.ID_PROGRAMA="&vId
sql = sql&" and HISTORICO_CURSOS.ESTADO=0"

DATOS.Open sql,oConn

Response.Write("<?xml version='1.0' encoding='utf-8' ?>"&chr(13))
Response.Write("<curriculo>"&chr(13)) 

WHILE NOT DATOS.EOF
		Response.Write("<row>"&chr(13))
		Response.Write("<ID>"&DATOS("ID_EMPRESA")&"</ID>"&chr(13))
		Response.Write("<NOMBRES>"&DATOS("R_SOCIAL")&"</NOMBRES>"&chr(13))
		Response.Write("</row>"&chr(13))
	DATOS.MoveNext
WEND
Response.Write("</curriculo>") 
%>
