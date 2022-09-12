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

vIdEmpresa=Request("id_empresa")
vIdPrograma=Request("id_programa")

sql = "select TRABAJADOR.NOMBRES,TRABAJADOR.ID_TRABAJADOR from HISTORICO_CURSOS "
sql = sql&" inner join TRABAJADOR on TRABAJADOR.ID_TRABAJADOR=HISTORICO_CURSOS.ID_TRABAJADOR "
sql = sql&" where HISTORICO_CURSOS.ESTADO=0 and HISTORICO_CURSOS.ID_PROGRAMA="&vIdPrograma&" and HISTORICO_CURSOS.ID_EMPRESA="&vIdEmpresa

DATOS.Open sql,oConn

Response.Write("<?xml version='1.0' encoding='utf-8' ?>"&chr(13))
Response.Write("<curriculo>"&chr(13)) 

WHILE NOT DATOS.EOF
		Response.Write("<row>"&chr(13))
		Response.Write("<ID>"&DATOS("ID_TRABAJADOR")&"</ID>"&chr(13))
		Response.Write("<NOMBRES>"&DATOS("NOMBRES")&"</NOMBRES>"&chr(13))
		Response.Write("</row>"&chr(13))
	DATOS.MoveNext
WEND
Response.Write("</curriculo>") 
%>
