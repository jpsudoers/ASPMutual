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

sql = "select id_empresa, ID_CONDICION_COMERCIAL , ISNULL(con_hes,0) con_hes, ISNULL(con_migo,0) con_migo "
sql = sql & " from EMPRESA_CONDICION_COMERCIAL "
sql = sql & " where ID_CONDICION_COMERCIAL = 0 " 


if(Request("Empresa")<>"")then
	sql = sql & " and id_empresa = " & request("Empresa")
end if

'response.write sql

DATOS.Open sql,oConn

Response.Write("<?xml version='1.0' encoding='utf-8' ?>"&chr(13))
Response.Write("<curriculo>"&chr(13)) 

if(DATOS("con_hes")="1") then 
		Response.Write("<row>"&chr(13))
		Response.Write("<ID_TIPO_REFERENCIA>1</ID_TIPO_REFERENCIA>"&chr(13))
		Response.Write("<TIPO_REFERENCIA>HES</TIPO_REFERENCIA>"&chr(13))
		Response.Write("</row>"&chr(13))
end if

if(DATOS("con_migo")="1") then 
		Response.Write("<row>"&chr(13))
		Response.Write("<ID_TIPO_REFERENCIA>2</ID_TIPO_REFERENCIA>"&chr(13))
		Response.Write("<TIPO_REFERENCIA>MIGO</TIPO_REFERENCIA>"&chr(13))
		Response.Write("</row>"&chr(13))
end if

Response.Write("</curriculo>") 
%>
