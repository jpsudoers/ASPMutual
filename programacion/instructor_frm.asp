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

'sql = "select InsRel.ID_INSTRUCTOR,dbo.MayMinTexto (InsRel.NOMBRES+' '+InsRel.A_PATERNO+' '+InsRel.A_MATERNO) as instructor " 
'sql = sql&" from INSTRUCTOR_RELATOR InsRel "
'sql = sql&" where InsRel.ID_INSTRUCTOR not in (select bq.id_relator from bloque_programacion bq where "
'sql = sql&" bq.id_relator<>'"&Request("idRelator")&"' and bq.id_programa='"&Request("progId")&"') "
'sql = sql&" and InsRel.ID_INSTRUCTOR not in (select bq.id_rel_seg from bloque_programacion bq where " 
'sql = sql&" bq.id_programa='"&Request("progId")&"' and bq.id_rel_seg is not NULL) "
'sql = sql&" and InsRel.ESTADO=1 order by NOMBRES asc "

sql = "select InsRel.ID_INSTRUCTOR,"
sql = sql&"dbo.MayMinTexto (InsRel.NOMBRES+' '+InsRel.A_PATERNO+' '+InsRel.A_MATERNO) as instructor " 
sql = sql&" from INSTRUCTOR_RELATOR InsRel "
'sql = sql&" where InsRel.ID_INSTRUCTOR not in (select bq.id_relator from bloque_programacion bq where "
'sql = sql&" bq.id_relator<>'"&Request("idRelator")&"' and bq.id_programa='"&Request("progId")&"') "
'sql = sql&" and InsRel.ID_INSTRUCTOR not in (select bq.id_rel_seg from bloque_programacion bq where " 
'sql = sql&" bq.id_rel_seg<>'"&Request("Relatorseg")&"' and bq.id_programa='"&Request("progId")&"' and bq.id_rel_seg is not NULL) "
sql = sql&" where /*and*/ InsRel.ESTADO=1 order by NOMBRES asc "

'response.Write(sql)
'response.End()

DATOS.Open sql,oConn

Response.Write("<?xml version='1.0' encoding='utf-8' ?>"&chr(13))
Response.Write("<instructor>"&chr(13)) 

WHILE NOT DATOS.EOF
		Response.Write("<row>"&chr(13))
		Response.Write("<ID_INSTRUCTOR>"&DATOS("ID_INSTRUCTOR")&"</ID_INSTRUCTOR>"&chr(13))
		Response.Write("<NOMBRES>"&DATOS("instructor")&"</NOMBRES>"&chr(13))
		Response.Write("</row>"&chr(13))
	DATOS.MoveNext
WEND
Response.Write("</instructor>") 
%>
