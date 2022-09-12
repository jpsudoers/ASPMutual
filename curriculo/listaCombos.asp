<!--#include file="../cnn_string.asp"-->
<%
Response.ContentType = "text/xml"
Response.AddHeader "Cache-control", "private"
Response.AddHeader "Expires", "-1"
Response.CodePage = 65001
Response.CharSet = "utf-8"
%>
<%
'Dim DATOS
'Dim oConn
'SET oConn = Server.CreateObject("ADODB.Connection")
'oConn.Open("Provider=SQLOLEDB; User ID=sa;Password=SCL.2013.2013;data source=.\SQLEXPRESS;Initial Catalog=dbmas")
'oConn.Open(MM_cnn_STRING)
'Set DATOS = Server.CreateObject("ADODB.RecordSet")
'DATOS.CursorType=3

'sql = "select * from CURRICULO where estado=1 order by NOMBRE_CURSO asc"

'DATOS.Open sql,oConn

Response.Write("<?xml version='1.0' encoding='utf-8' ?>"&chr(13))
Response.Write("<combos>"&chr(13)) 

'WHILE NOT DATOS.EOF
if(Request("t")="1")then
	For i=1 to 5
			Response.Write("<row>"&chr(13))
			Response.Write("<ID>"&i&"</ID>"&chr(13))
			Response.Write("<NOMBRES>"&i&"</NOMBRES>"&chr(13))
			Response.Write("</row>"&chr(13))
	Next 	
elseif(Request("t")="2")then
		For i=1 to 4
			Response.Write("<row>"&chr(13))
			Response.Write("<ID>Bloque "&i&"</ID>"&chr(13))
			Response.Write("<NOMBRES>Bloque "&i&"</NOMBRES>"&chr(13))
			Response.Write("</row>"&chr(13))
		Next 		
elseif(Request("t")="3")then
		For i=1 to 8
			Response.Write("<row>"&chr(13))
			Response.Write("<ID>"&i&":00</ID>"&chr(13))
			Response.Write("<NOMBRES>"&i&":00</NOMBRES>"&chr(13))
			Response.Write("</row>"&chr(13))		
			'if((cint(i) mod 2) = 0)then
				Response.Write("<row>"&chr(13))
				Response.Write("<ID>"&i&":30</ID>"&chr(13))
				Response.Write("<NOMBRES>"&i&":30</NOMBRES>"&chr(13))
				Response.Write("</row>"&chr(13))
			'end if
		Next 			
end if
	'DATOS.MoveNext
'WEND
Response.Write("</combos>") 
%>
