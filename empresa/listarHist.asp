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
'oConn.Open("Provider=SQLOLEDB; User ID=sa;Password=SCL.2013.2013;data source=.\SQLEXPRESS;Initial Catalog=dbarauco")
oConn.Open(MM_cnn_STRING)
Set DATOS = Server.CreateObject("ADODB.RecordSet")
DATOS.CursorType=3

vemp=Request("e")

sql = "select EST=(CASE eb.ESTADO When 0 then 'Bloqueo' else 'Desbloqueo' end),eb.FECHA_ESTADO,"&_
	  "nom_usuario=u.NOMBRES+' '+u.A_PATERNO,eb.DESCRIPCION from EMPRESAS_BLOQUEOS eb "&_
	  " inner join USUARIOS u on u.ID_USUARIO=eb.ID_USUARIO "&_
	  " where eb.ID_EMPRESA='"&vemp&"' /*and eb.ID_CLIENTE=1*/ "&_
	  " order by eb.FECHA_ESTADO desc "

DATOS.Open sql,oConn

Response.Write("<?xml version='1.0' encoding='utf-8' ?>"&chr(13))
Response.Write("<rows>"&chr(13)) 

fila=0
WHILE NOT DATOS.EOF
		Response.Write("<row id="""&fila&""">"&chr(13))
		Response.Write("<cell><![CDATA["&DATOS("EST")&"]]></cell>"&chr(13))
		Response.Write("<cell><![CDATA["&DATOS("FECHA_ESTADO")&"]]></cell>"&chr(13))
		Response.Write("<cell><![CDATA["&DATOS("nom_usuario")&"]]></cell>"&chr(13))
		Response.Write("<cell><![CDATA["&DATOS("DESCRIPCION")&"]]></cell>"&chr(13))
		Response.Write("</row>"&chr(13))
		fila=fila+1
		DATOS.MoveNext
WEND
Response.Write("</rows>") 
%>
