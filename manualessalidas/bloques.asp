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

sql = "select BP.id_bloque,dbo.MayMinTexto(IR.NOMBRES+' '+IR.A_PATERNO+' '+IR.A_MATERNO) as RELATOR from bloque_programacion BP "
sql = sql&" inner join INSTRUCTOR_RELATOR IR ON IR.ID_INSTRUCTOR=BP.id_relator "
sql = sql&" where BP.id_programa in (select paux.ID_PROGRAMA from PROGRAMA paux "
sql = sql&" where CONVERT(date,paux.FECHA_INICIO_)=CONVERT(date,'"&Request("id_programa")&"') and "
sql = sql&"paux.ID_MUTUAL='"&Request("curso")&"' and paux.ESTADO=1 and paux.VIGENCIA=1) "
sql = sql&" ORDER by IR.NOMBRES ASC "

'response.Write(sql)
'response.End()

DATOS.Open sql,oConn

Response.Write("<?xml version='1.0' encoding='utf-8' ?>"&chr(13))
Response.Write("<curriculo>"&chr(13)) 

WHILE NOT DATOS.EOF
		Response.Write("<row>"&chr(13))
		Response.Write("<ID_BLOQUE>"&DATOS("id_bloque")&"</ID_BLOQUE>"&chr(13))
		Response.Write("<RELATOR>"&DATOS("RELATOR")&"</RELATOR>"&chr(13))
		Response.Write("</row>"&chr(13))
	DATOS.MoveNext
WEND
Response.Write("</curriculo>") 
%>
