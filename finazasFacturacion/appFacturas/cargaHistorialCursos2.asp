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

  sql = "select a.ID_EMPRESA,a.ID_OTIC,"&_
		" ult_fac_emp=ISNULL((select top 1 f.ESTADO from facturas f "&_
		" where f.ID_AUTORIZACION=a.ID_AUTORIZACION "&_ 
		" and f.ID_EMPRESA=a.ID_EMPRESA order by f.ID_FACTURA desc),2),"&_
		" ult_fac_otic=ISNULL((select top 1 f.ESTADO from facturas f "&_
		" where f.ID_AUTORIZACION=a.ID_AUTORIZACION "&_
		" and f.ID_EMPRESA=a.ID_OTIC order by f.ID_FACTURA desc),2) "&_
		" from AUTORIZACION a "&_
		" where a.ID_AUTORIZACION='"&Request("idAuto")&"'"

'response.Write(sql)
'response.End()

DATOS.Open sql,oConn

Response.Write("<?xml version='1.0' encoding='utf-8' ?>"&chr(13))
Response.Write("<rows>"&chr(13)) 
Response.Write("<records>"&DATOS.RecordCount&"</records>"&chr(13))

fila=1
'WHILE NOT DATOS.EOF
	Response.Write("<row id="""&fila&""">"&chr(13))
	
if( (DATOS("ult_fac_emp")="2" and DATOS("ult_fac_otic")="2") or (DATOS("ult_fac_emp")="0" and DATOS("ult_fac_otic")="0") or (DATOS("ult_fac_emp")="0" and DATOS("ult_fac_otic")="2") or (DATOS("ult_fac_emp")="2" and DATOS("ult_fac_otic")="0"))then	
			Response.Write("<Empresa>completa</Empresa>"&chr(13))
	else
			if(DATOS("ult_fac_emp")="0" and DATOS("ult_fac_otic")="1")then
				Response.Write("<Empresa>solo empresa</Empresa>"&chr(13))
			end if
			
			if(DATOS("ult_fac_emp")="1" and DATOS("ult_fac_otic")="0")then
				Response.Write("<Empresa>solo otic</Empresa>"&chr(13))		
			end if
	end if

	Response.Write("</row>"&chr(13))

	'fila=fila+1
	'DATOS.MoveNext
'WEND
Response.Write("</rows>") 
%>