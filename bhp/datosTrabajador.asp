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

vID = Request("id")

sql = "select ID_TRABAJADOR,RUT,UPPER(NOMBRES) as NOMBRES from TRABAJADOR where ID_TRABAJADOR='"&vID&"'"

DATOS.Open sql,oConn

Response.Write("<?xml version='1.0' encoding='utf-8' ?>"&chr(13))
Response.Write("<DATOS>"&chr(13)) 

WHILE NOT DATOS.EOF
		Response.Write("<row>"&chr(13))
		Response.Write("<IDTRABAJADOR>"&DATOS("ID_TRABAJADOR")&"</IDTRABAJADOR>"&chr(13))
		Response.Write("<NOMBRE>"&DATOS("NOMBRES")&"</NOMBRE>"&chr(13))
		Response.Write("<RUT>"&DATOS("RUT")&"</RUT>"&chr(13))
		Response.Write("</row>"&chr(13))
	DATOS.MoveNext
WEND
Response.Write("</DATOS>") 
%>
