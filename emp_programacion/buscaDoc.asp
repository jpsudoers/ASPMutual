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

sql = "IF NOT EXISTS (select * from preinscripciones p where p.id_empresa='"&Request("id_empresa")&"'"
sql = sql&" and p.tipo_compromiso='"&Request("tipo_doc")&"' and p.numero_compromiso='"&Request("n_doc")&"' and p.estado=1)"
sql = sql&" and NOT EXISTS (select * from AUTORIZACION a where a.ID_EMPRESA='"&Request("id_empresa")&"'"
sql = sql&" and a.TIPO_DOC='"&Request("tipo_doc")&"' and a.ORDEN_COMPRA='"&Request("n_doc")&"') BEGIN "
sql = sql&" select '1' as 'resultado' end else begin  select '' as 'resultado' END"

DATOS.Open sql,oConn

Response.Write("<?xml version='1.0' encoding='utf-8' ?>"&chr(13))
Response.Write("<DATOS>"&chr(13)) 
Response.Write("<row>"&chr(13))
Response.Write("<Existe>"&DATOS("resultado")&"</Existe>"&chr(13))
Response.Write("</row>"&chr(13))
Response.Write("</DATOS>") 
%>
