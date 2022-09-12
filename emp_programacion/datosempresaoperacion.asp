<!--#include file="../conexion.asp"-->
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

vID = Request("id")

sql = "SELECT EMPRESAS.ID_EMPRESA as empresa,EMPRESAS.R_SOCIAL as empresarsocial,EMPRESAS.RUT as empresarut,EMPRESAS.ID_OTIC,EMPRESAS.TIPO,EST=ISNULL((select top 1 eb.ESTADO from EMPRESAS_BLOQUEOS eb where eb.ID_EMPRESA=EMPRESAS.ID_EMPRESA order by eb.ID_BLOQUEO_EMPRESA desc),1) "
sql = sql&" FROM EMPRESAS where EMPRESAS.ID_EMPRESA='"&vID&"'"

DATOS.Open sql,oConn

Response.Write("<?xml version='1.0' encoding='utf-8' ?>"&chr(13))
Response.Write("<DATOS>"&chr(13)) 

WHILE NOT DATOS.EOF
		Response.Write("<row>"&chr(13))
		Response.Write("<IDEMPRESA>"&DATOS("empresa")&"</IDEMPRESA>"&chr(13))
		Response.Write("<RUT>"&DATOS("empresarut")&"</RUT>"&chr(13))
		Response.Write("<RSOCIAL>"&DATOS("empresarsocial")&"</RSOCIAL>"&chr(13))
		Response.Write("<IDOTIC>"&DATOS("ID_OTIC")&"</IDOTIC>"&chr(13))

		sqlOtic = "SELECT EMPRESAS.R_SOCIAL as empresarsocial,EMPRESAS.RUT as empresarut"
		sqlOtic = sqlOtic&" FROM EMPRESAS where EMPRESAS.ID_EMPRESA='"&DATOS("ID_OTIC")&"'"
		
		set rsOtic = conn.execute (sqlOtic)

		   if not rsOtic.eof and not rsOtic.bof then 
			Response.Write("<RSOCIALOTIC>"&rsOtic("empresarsocial")&"</RSOCIALOTIC>"&chr(13))
			Response.Write("<RUTOTIC>"&replace(FormatNumber(mid(rsOtic("empresarut"), 1,len(rsOtic("empresarut"))-2),0)&mid(rsOtic("empresarut"), len(rsOtic("empresarut"))-1,len(rsOtic("empresarut"))),",",".")&"</RUTOTIC>"&chr(13))
		   else
			Response.Write("<RSOCIALOTIC>0</RSOCIALOTIC>"&chr(13))
			Response.Write("<RUTOTIC>0</RUTOTIC>"&chr(13))
		   end if
		
		Response.Write("<TIPO>"&DATOS("TIPO")&"</TIPO>"&chr(13))
		Response.Write("<BLOQUEO>"&DATOS("EST")&"</BLOQUEO>"&chr(13))		
		Response.Write("</row>"&chr(13))
	DATOS.MoveNext
WEND
Response.Write("</DATOS>") 
%>
