<!--#include file="../cnn_string.asp"-->
<%
Response.ContentType = "text/xml"
Response.AddHeader "Cache-control", "private"
Response.AddHeader "Expires", "-1"
Response.CodePage = 65001
Response.CharSet = "utf-8"
%>
<!--#include file="../conexion.asp"-->
<%
Dim DATOS
Dim oConn
SET oConn = Server.CreateObject("ADODB.Connection")
'oConn.Open("Provider=SQLOLEDB; User ID=sa;Password=SCL.2013.2013;data source=.\SQLEXPRESS;Initial Catalog=dbmas")
oConn.Open(MM_cnn_STRING)
Set DATOS = Server.CreateObject("ADODB.RecordSet")
DATOS.CursorType=3

vsol=Request("id")

sql = "select SOLICITUD_TRABAJADOR.id_trabajador from SOLICITUD_TRABAJADOR "
sql = sql&" inner join SOLICITUD on SOLICITUD.id_solicitud=SOLICITUD_TRABAJADOR.id_solicitud "
sql = sql&" where SOLICITUD.id_solicitud="&vsol
sql = sql&" order by SOLICITUD_TRABAJADOR.id_trabajador asc "

DATOS.Open sql,oConn

Response.Write("<?xml version='1.0' encoding='utf-8' ?>"&chr(13))
Response.Write("<rows>"&chr(13)) 

fila=0
WHILE NOT DATOS.EOF
		Response.Write("<row id="""&fila&""">"&chr(13))
		
		sql2 = "select TRABAJADOR.RUT,TRABAJADOR.NOMBRES,TRABAJADOR.CARGO_EMPRESA,TRABAJADOR.ESCOLARIDAD "
		sql2 = sql2&" from TRABAJADOR "
		sql2 = sql2&" where TRABAJADOR.ID_TRABAJADOR="&DATOS("id_trabajador")
		
		set rsTrab = conn.execute (sql2)
		
		Response.Write("<cell><![CDATA["&rsTrab("RUT")&"]]></cell>"&chr(13))
		Response.Write("<cell><![CDATA["&rsTrab("NOMBRES")&"]]></cell>"&chr(13))
		Response.Write("<cell><![CDATA["&rsTrab("CARGO_EMPRESA")&"]]></cell>"&chr(13))
		
		if(rsTrab("ESCOLARIDAD")=0)then
		Response.Write("<cell><![CDATA[Sin Escolaridad]]></cell>"&chr(13))
		end if
		
		if(rsTrab("ESCOLARIDAD")=1)then
		Response.Write("<cell><![CDATA[Básica Incompleta]]></cell>"&chr(13))
		end if
		
		if(rsTrab("ESCOLARIDAD")=2)then
		Response.Write("<cell><![CDATA[Básica Completa]]></cell>"&chr(13))
		end if


		if(rsTrab("ESCOLARIDAD")=3)then
		Response.Write("<cell><![CDATA[Media Incompleta]]></cell>"&chr(13))
		end if

		if(rsTrab("ESCOLARIDAD")=4)then
		Response.Write("<cell><![CDATA[Media Completa]]></cell>"&chr(13))
		end if
		
		if(rsTrab("ESCOLARIDAD")=5)then
		Response.Write("<cell><![CDATA[Superior Técnica Incompleta]]></cell>"&chr(13))
		end if

		if(rsTrab("ESCOLARIDAD")=6)then
		Response.Write("<cell><![CDATA[Superior Técnica Profesional Completa]]></cell>"&chr(13))
		end if

		if(rsTrab("ESCOLARIDAD")=7)then
		Response.Write("<cell><![CDATA[Universitaria Incompleta]]></cell>"&chr(13))
		end if
		
		if(rsTrab("ESCOLARIDAD")=8)then
		Response.Write("<cell><![CDATA[Universitaria Completa]]></cell>"&chr(13))
		end if
		
		Response.Write("</row>"&chr(13))
		fila=fila+1
		DATOS.MoveNext
WEND
Response.Write("<row id="""&fila&""">"&chr(13))
Response.Write("<cell><![CDATA[]]></cell>"&chr(13))
Response.Write("<cell><![CDATA[]]></cell>"&chr(13))
Response.Write("<cell><![CDATA[]]></cell>"&chr(13))
Response.Write("<cell><![CDATA[]]></cell>"&chr(13))
Response.Write("</row>"&chr(13))
Response.Write("</rows>") 
%>
