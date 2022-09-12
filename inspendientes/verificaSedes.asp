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

sql = "select COUNT(HISTORICO_CURSOS.SEDE)as totalSede from HISTORICO_CURSOS "
sql = sql&" inner join SEDES on SEDES.ID_SEDE=HISTORICO_CURSOS.SEDE "
sql = sql&" where HISTORICO_CURSOS.RELATOR="&Request("relator")&" and HISTORICO_CURSOS.ID_PROGRAMA="&Request("prog")&""

DATOS.Open sql,oConn

Response.Write("<?xml version='1.0' encoding='utf-8' ?>"&chr(13))
Response.Write("<sede>"&chr(13)) 

WHILE NOT DATOS.EOF
		Response.Write("<row>"&chr(13))
		Response.Write("<totalSede>"&DATOS("totalSede")&"</totalSede>"&chr(13))
		Response.Write("</row>"&chr(13))
	DATOS.MoveNext
WEND
Response.Write("</sede>") 
%>
