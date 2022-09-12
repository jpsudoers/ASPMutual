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

Dim vrelator
vrelator = "0"
If (Request("relator") <> "") Then 
  vrelator = Request("relator")
End If

Dim vsede
vsede = "0"
If (Request("sede") <> "") Then 
  vsede = Request("sede")
End If


sql = "select COUNT(HISTORICO_CURSOS.ID_HISTORICO_CURSO) as inscritos from HISTORICO_CURSOS "
sql = sql&" inner join TRABAJADOR on TRABAJADOR.ID_TRABAJADOR=HISTORICO_CURSOS.ID_TRABAJADOR "
sql = sql&" where HISTORICO_CURSOS.ID_PROGRAMA="&Request("programa")&" and HISTORICO_CURSOS.RELATOR="&vrelator
sql = sql&" and HISTORICO_CURSOS.SEDE="&vsede

DATOS.Open sql,oConn

Response.Write("<?xml version='1.0' encoding='utf-8' ?>"&chr(13))
Response.Write("<totales>"&chr(13)) 

WHILE NOT DATOS.EOF
			Response.Write("<row>"&chr(13))
			Response.Write("<total>"&DATOS("inscritos")&"</total>"&chr(13))
			Response.Write("</row>"&chr(13))
	DATOS.MoveNext
WEND
Response.Write("</totales>") 
%>
