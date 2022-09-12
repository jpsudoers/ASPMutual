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

vID = Request("id_programa")

sql = "select PROGRAMA.ID_PROGRAMA, "
sql = sql&"INSTRUCTOR_RELATOR.NOMBRES+' '+INSTRUCTOR_RELATOR.A_PATERNO+' '+INSTRUCTOR_RELATOR.A_MATERNO as nombre, "
sql = sql&"INSTRUCTOR_RELATOR.RUT, SEDES.NOMBRE as sede, SEDES.DIRECCION from PROGRAMA  "
sql = sql&" inner join INSTRUCTOR_RELATOR on INSTRUCTOR_RELATOR.ID_INSTRUCTOR=PROGRAMA.ID_INSTRUCTOR "
sql = sql&" inner join SEDES on SEDES.ID_SEDE=PROGRAMA.ID_SEDE "
sql = sql&" where PROGRAMA.ID_PROGRAMA='"&vID&"'"

DATOS.Open sql,oConn

Response.Write("<?xml version='1.0' encoding='utf-8' ?>"&chr(13))
Response.Write("<DATOS>"&chr(13)) 

WHILE NOT DATOS.EOF
		Response.Write("<row>"&chr(13))
		Response.Write("<IDPROGRAMA>"&DATOS("ID_PROGRAMA")&"</IDPROGRAMA>"&chr(13))
		Response.Write("<instructor>"&DATOS("nombre")&"</instructor>"&chr(13))
		Response.Write("<RUT>"&DATOS("RUT")&"</RUT>"&chr(13))
		Response.Write("<SEDE>"&DATOS("sede")&"</SEDE>"&chr(13))
		Response.Write("<DIRECCION>"&DATOS("DIRECCION")&"</DIRECCION>"&chr(13))
		Response.Write("</row>"&chr(13))
	DATOS.MoveNext
WEND
Response.Write("</DATOS>") 
%>
