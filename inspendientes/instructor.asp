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

sql = "select INSTRUCTOR_RELATOR.ID_INSTRUCTOR,"
sql = sql&"INSTRUCTOR_RELATOR.NOMBRES+' '+INSTRUCTOR_RELATOR.A_PATERNO+' '+INSTRUCTOR_RELATOR.A_MATERNO as instructor, "
sql = sql&"(select COUNT(HISTORICO_CURSOS.ID_HISTORICO_CURSO) from HISTORICO_CURSOS " 
sql = sql&"where HISTORICO_CURSOS.RELATOR=INSTRUCTOR_RELATOR.ID_INSTRUCTOR "
sql = sql&"and HISTORICO_CURSOS.ID_PROGRAMA="&Request("instructor")&") as inscritos "
sql = sql&" from INSTRUCTOR_RELATOR "
sql = sql&" where INSTRUCTOR_RELATOR.ESTADO=1 "&Request("id_relator")&" order by NOMBRES asc "

DATOS.Open sql,oConn

Response.Write("<?xml version='1.0' encoding='utf-8' ?>"&chr(13))
Response.Write("<instructor>"&chr(13)) 

WHILE NOT DATOS.EOF
		if(cdbl(DATOS("inscritos"))<30)then
			Response.Write("<row>"&chr(13))
			Response.Write("<ID_INSTRUCTOR>"&DATOS("ID_INSTRUCTOR")&"</ID_INSTRUCTOR>"&chr(13))
			Response.Write("<NOMBRES>"&DATOS("instructor")&"</NOMBRES>"&chr(13))
			Response.Write("</row>"&chr(13))
		end if
	DATOS.MoveNext
WEND
Response.Write("</instructor>") 
%>
