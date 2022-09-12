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

'sql = "select COUNT(*) as TPreinsTrab from PREINSCRIPCION_TRABAJADOR preins "
'sql = sql&" where preins.id_trabajador='"&Request("trab")&"' and preins.preinscripcionTemp='"&Request("id")&"'"

sql = "select ins_auto=(select count(*) from HISTORICO_CURSOS hc "&_
" where hc.ID_TRABAJADOR in ('"&Request("trab")&"') and hc.ID_PROGRAMA in (select p.ID_PROGRAMA from programa p "&_
" where p.FECHA_INICIO_=convert(date, '"&Request("fecha")&"'))), "&_
" ins_pre=(select COUNT(*) from PREINSCRIPCION_TRABAJADOR pt "&_
" inner join preinscripciones p on p.id_preinscripcion=pt.id_preinscripcion "&_
" where pt.id_trabajador in ('"&Request("trab")&"') and p.id_programacion in (select p2.ID_PROGRAMA from programa p2 "&_
" where p2.FECHA_INICIO_=convert(date, '"&Request("fecha")&"')) and p.estado=1)"

DATOS.Open sql,oConn

Response.Write("<?xml version='1.0' encoding='utf-8' ?>"&chr(13))
Response.Write("<DATOS>"&chr(13)) 

WHILE NOT DATOS.EOF
		Response.Write("<row>"&chr(13))
		if(cint(DATOS("ins_auto"))>0 or cint(DATOS("ins_pre"))>0)then
			Response.Write("<BusTrabCur></BusTrabCur>"&chr(13))
		else
			Response.Write("<BusTrabCur>1</BusTrabCur>"&chr(13))
		end if
		Response.Write("</row>"&chr(13))
	DATOS.MoveNext
WEND
Response.Write("</DATOS>") 
%>
