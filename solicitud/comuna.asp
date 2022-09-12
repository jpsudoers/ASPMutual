<%Response.CodePage = 65001%>


<!--#include file="../cnn_string.asp"-->
<%


SET oConn = Server.CreateObject("ADODB.Connection")
oConn.Open(MM_cnn_STRING)
Set DATOS = Server.CreateObject("ADODB.RecordSet")
DATOS.CursorType=3
sql = "select upper(nombre) as nombre from comuna ORDER BY NOMBRE"
DATOS.Open sql,oConn

	Response.Write("<?xml version='1.0' encoding='utf-8' ?>"&chr(13))
	Response.Write("<comuna>"&chr(13)) 
	
	WHILE NOT DATOS.EOF
			Response.Write("<row>"&chr(13))
			Response.Write("<ID>"&DATOS("nombre")&"</ID>"&chr(13))
			Response.Write("<DESC>"&DATOS("nombre")&"</DESC>"&chr(13))
			Response.Write("</row>"&chr(13))
		DATOS.MoveNext
	WEND
	Response.Write("</comuna>") 
%>