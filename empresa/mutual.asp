<!--#include file="../cnn_string.asp"-->
<%
'Response.ContentType = "text/xml"
Response.AddHeader "Cache-control", "private"
Response.AddHeader "Expires", "-1"
Response.CodePage = 65001
Response.CharSet = "utf-8"
%>
<%
Dim DATOS
Dim oConn
Dim MUTUAL
SET oConn = Server.CreateObject("ADODB.Connection")
'oConn.Open("Provider=SQLOLEDB; User ID=sa;Password=SCL.2013.2013;data source=.\SQLEXPRESS;Initial Catalog=dbmas")
oConn.Open(MM_cnn_STRING)
MUTUAL=MM_cnn_STRING
Set DATOS = Server.CreateObject("ADODB.RecordSet")
DATOS.CursorType=3

sql = "select * from EMPRESAS where estado=1 and tipo=3 order by R_SOCIAL asc"

DATOS.Open sql,oConn

Response.Write("<?xml version='1.0' encoding='utf-8' ?>"&chr(13))
Response.Write("<mutual>"&chr(13)) 
Response.Write("<ID_MUTUAL_><![CDATA["&MUTUAL&"]]></ID_MUTUAL_>"&chr(13))
WHILE NOT DATOS.EOF
		Response.Write("<row>"&chr(13))
		
		Response.Write("<ID_MUTUAL>"&DATOS("ID_EMPRESA")&"</ID_MUTUAL>"&chr(13))
		Response.Write("<NOMBRES>"&DATOS("R_SOCIAL")&"</NOMBRES>"&chr(13))
		Response.Write("</row>"&chr(13))
	DATOS.MoveNext
WEND
Response.Write("</mutual>") 
%>
