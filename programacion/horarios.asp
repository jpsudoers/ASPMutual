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
oConn.Open(MM_cnn_STRING)
Set DATOS = Server.CreateObject("ADODB.RecordSet")
DATOS.CursorType=3

sql = "select hb.ID_HORARIO,hb.HORARIO from horario_bloques hb "&_
      " where (hb.ID_HORARIO>="&Request("i")&" and hb.ID_HORARIO<="&Request("t")&") or hb.ID_HORARIO in (1010,1012)"&_
      " order by hb.HORARIO asc"

DATOS.Open sql,oConn

Response.Write("<?xml version='1.0' encoding='utf-8' ?>"&chr(13))
Response.Write("<horario>"&chr(13)) 

WHILE NOT DATOS.EOF
		Response.Write("<row>"&chr(13))
		Response.Write("<ID>"&DATOS("ID_HORARIO")&"</ID>"&chr(13))
		Response.Write("<NOMBRE>"&DATOS("HORARIO")&"</NOMBRE>"&chr(13))
		Response.Write("</row>"&chr(13))
	DATOS.MoveNext
WEND
Response.Write("</horario>") 
%>
