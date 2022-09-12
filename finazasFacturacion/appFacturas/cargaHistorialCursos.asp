<!--#include file="../../cnn_string.asp"-->
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

sql = "select c.ID_MUTUAL,C.CODIGO,dbo.MayMinTexto(C.NOMBRE_CURSO) as nombre_curso,(select COUNT(*) from AUTORIZACION A inner join PROGRAMA p on p.ID_PROGRAMA=A.ID_PROGRAMA where A.ESTADO=0 and A.FACTURADO=1 AND P.ID_MUTUAL=C.ID_MUTUAL AND YEAR(P.FECHA_INICIO_)in (2007,2012,2013,2014,2015,2016,2017,2018,2019,2020,2021,2022) /*AND A.FECHA_rev_cierre<=convert(datetime, '23-06-2016 18:00:00')*/) AS TOTAL   from CURRICULO C where C.ESTADO IN (0,1) order by TOTAL desc"

DATOS.Open sql,oConn

Response.Write("<?xml version='1.0' encoding='utf-8' ?>"&chr(13))
Response.Write("<rows>"&chr(13)) 
Response.Write("<records>"&DATOS.RecordCount&"</records>"&chr(13))

fila=0
WHILE NOT DATOS.EOF
	Response.Write("<row id="""&fila&""">"&chr(13))
	Response.Write("<CODIGO>"&DATOS("CODIGO")&"</CODIGO>"&chr(13))
	Response.Write("<nombre_curso>"&DATOS("nombre_curso")&"</nombre_curso>"&chr(13))
	Response.Write("<TOTAL>"&DATOS("TOTAL")&"</TOTAL>"&chr(13))
	Response.Write("<ID_MUTUAL>"&DATOS("ID_MUTUAL")&"</ID_MUTUAL>"&chr(13))
	Response.Write("</row>"&chr(13))

	fila=fila+1
	DATOS.MoveNext
WEND
Response.Write("</rows>") 
%>