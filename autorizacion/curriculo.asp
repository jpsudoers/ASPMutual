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

sql = "SELECT EMPRESAS.ID_EMPRESA as empresa,EMPRESAS.R_SOCIAL as empresarsocial,EMPRESAS.RUT as empresarut, "
sql = sql&"OTIC.ID_OTIC as otic ,otic.R_SOCIAL as oticrsoacial,OTIC.RUT as oticrut FROM EMPRESAS "
sql = sql&" inner join OTIC on OTIC.ID_OTIC=EMPRESAS.ID_OTIC "
sql = sql&" where EMPRESAS.ID_EMPRESA='"&vID&"'"

DATOS.Open sql,oConn

Response.Write("<?xml version='1.0' encoding='utf-8' ?>"&chr(13))
Response.Write("<DATOS>"&chr(13)) 

WHILE NOT DATOS.EOF
		Response.Write("<row>"&chr(13))
		Response.Write("<IDEMPRESA>"&DATOS("empresa")&"</IDEMPRESA>"&chr(13))
		Response.Write("<RSOCIAL>"&DATOS("empresarsocial")&"</RSOCIAL>"&chr(13))
		Response.Write("<RUT>"&DATOS("empresarut")&"</RUT>"&chr(13))
		Response.Write("<IDOTIC>"&DATOS("otic")&"</IDOTIC>"&chr(13))
		Response.Write("<RSOCIALOTIC>"&DATOS("oticrsoacial")&"</RSOCIALOTIC>"&chr(13))
		Response.Write("<RUTOTIC>"&DATOS("oticrut")&"</RUTOTIC>"&chr(13))
		Response.Write("</row>"&chr(13))
	DATOS.MoveNext
WEND
Response.Write("</DATOS>") 
%>
