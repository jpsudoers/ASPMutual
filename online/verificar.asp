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

vcod = Request("c")

sql = "select hc.ID_PROGRAMA,hc.ID_EMPRESA,hc.ID_TRABAJADOR,hc.RELATOR from HISTORICO_CURSOS hc where hc.COD_AUTENFICACION='"&vcod&"'"

DATOS.Open sql,oConn

Response.Write("<?xml version='1.0' encoding='utf-8'?>"&chr(13))
Response.Write("<DATOS>"&chr(13)) 

Response.Write("<total>"&DATOS.RecordCount&"</total>"&chr(13))
WHILE NOT DATOS.EOF
		Response.Write("<row>"&chr(13))
		Response.Write("<p>"&DATOS("ID_PROGRAMA")&"</p>"&chr(13))
		Response.Write("<e>"&DATOS("ID_EMPRESA")&"</e>"&chr(13))
		Response.Write("<t>"&DATOS("ID_TRABAJADOR")&"</t>"&chr(13))
		Response.Write("<r>"&DATOS("RELATOR")&"</r>"&chr(13))
		Response.Write("</row>"&chr(13))
	DATOS.MoveNext
WEND
Response.Write("</DATOS>") 
%>
