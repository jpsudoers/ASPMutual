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

'sql = "select ID_USUARIO as usuario,ID_INSTRUCTOR as relUser,dbo.MayMinTexto(NOMBRES+' '+A_PATERNO+' '+A_MATERNO) as nom_user from USUARIOS where ESTADO=1 and RUT='" + Request("RUT") + "' and CONTRASEÑA= '" + Request("PASS") + "'"

sql = "IF CONVERT(date,GETDATE())>=CONVERT(date,dateadd(ms,-3,DATEADD(mm, DATEDIFF(m,0,(select ac.FECHA_APERTURA "&_
" from APERTURA_CIERRE ac where ac.ESTADO=1))+1, 0))) BEGIN "&_
" select (CONVERT(varchar,dateadd(ms,-3,DATEADD(mm, DATEDIFF(m,0,(select ac.FECHA_APERTURA "&_
" from APERTURA_CIERRE ac where ac.ESTADO=1))+1, 0)),105)) as fecha,(1) as estado end else BEGIN "&_
" select (CONVERT(varchar,GETDATE(),105)) as fecha,(0) as estado end"

'RESPONSE.Write(sql)
'RESPONSE.End()

DATOS.Open sql,oConn

Response.Write("<?xml version='1.0' encoding='utf-8' ?>"&chr(13))
Response.Write("<rows>"&chr(13)) 
Response.Write("<records>"&DATOS.RecordCount&"</records>"&chr(13))

fila=0
WHILE NOT DATOS.EOF
	Response.Write("<row id="""&fila&""">"&chr(13))
	Response.Write("<FECHA>"&DATOS("fecha")&"</FECHA>"&chr(13))
	Response.Write("<ESTADO>"&DATOS("estado")&"</ESTADO>"&chr(13))
	Response.Write("</row>"&chr(13))

	fila=fila+1
	DATOS.MoveNext
WEND
Response.Write("</rows>") 
%>