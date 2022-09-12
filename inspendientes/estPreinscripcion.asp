<!--#include file="../cnn_string.asp"-->
<!--#include file="../conexion.asp"-->
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

vID = Request("id")
sql="select p.estado,(case p.situacion when 1 then "&_
     "'Autorizada por '+(select  u.NOMBRES+' '+U.A_PATERNO+' '+U.A_MATERNO  from usuarios u "&_
     " where u.ID_USUARIO=p.situacion_encargado) "&_
	 " when 2 then 'Rechazada por '+(select  u.NOMBRES+' '+U.A_PATERNO+' '+U.A_MATERNO  from usuarios u "&_
     " where u.ID_USUARIO=p.situacion_encargado) "&_ 
     " when 3 then 'Eliminada por '+(select u.NOMBRES+' '+U.A_PATERNO+' '+U.A_MATERNO from usuarios u  "&_
     " where u.ID_USUARIO=p.situacion_encargado)  "&_
     " else '' END) AS COM_ESTADO from preinscripciones P where p.id_preinscripcion='"&vID&"'"

DATOS.Open sql,oConn

Response.Write("<?xml version='1.0' encoding='utf-8' ?>"&chr(13))
Response.Write("<DATOS>"&chr(13)) 

WHILE NOT DATOS.EOF
		Response.Write("<row>"&chr(13))
		Response.Write("<RESULTADO>"&DATOS("estado")&"</RESULTADO>"&chr(13))
		Response.Write("<COMENTARIO>"&DATOS("COM_ESTADO")&"</COMENTARIO>"&chr(13))		
		Response.Write("</row>"&chr(13))
	DATOS.MoveNext
WEND
Response.Write("</DATOS>") 
%>