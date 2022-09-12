<!--#include file="../cnn_string.asp"-->
<%
Response.ContentType = "text/xml"
Response.AddHeader "Cache-control", "private"
Response.AddHeader "Expires", "-1"
Response.CodePage = 65001
Response.CharSet = "utf-8"

pagina=1
if(Request("page")<>"")then pagina = CInt(Request("page"))

limite = 1
if(Request("rows")<>"")then limite = CInt(Request("rows"))
orden = "ASC"
if(Request("sord") <> "")then orden = Request("sord")
columna = "CURRICULO.NOMBRE_CURSO"
if(Request("sidx") <> "")then columna = Request("sidx")

node=0
if(request("nodeid")<>"")then node=cInt(request("nodeid"))

%>

<%
Dim DATOS
Dim oConn
SET oConn = Server.CreateObject("ADODB.Connection")
'oConn.Open("Provider=SQLOLEDB; User ID=sa;Password=SCL.2013.2013;data source=.\SQLEXPRESS;Initial Catalog=dbmas")
oConn.Open(MM_cnn_STRING)
Set DATOS = Server.CreateObject("ADODB.RecordSet")
DATOS.CursorType=3

sql = "SELECT SOLICITUD.id_solicitud, EMPRESAS.RUT, EMPRESAS.R_SOCIAL as empresa, MUTUALES.nomb_mutual, OTIC.R_SOCIAL, "
sql = sql&"CURRICULO.CODIGO,CURRICULO.NOMBRE_CURSO, CONVERT(VARCHAR(10),SOLICITUD.fecha, 105) as fecha, SOLICITUD.participantes "
sql = sql&" FROM SOLICITUD "
sql = sql&" INNER JOIN EMPRESAS ON EMPRESAS.ID_EMPRESA=SOLICITUD.id_empresa "
sql = sql&" INNER JOIN MUTUALES ON MUTUALES.Mutual_id =EMPRESAS.MUTUAL "
sql = sql&" INNER JOIN OTIC ON OTIC.ID_OTIC=EMPRESAS.ID_OTIC "
sql = sql&" INNER JOIN CURRICULO ON CURRICULO.ID_MUTUAL=SOLICITUD.id_mutual "
sql = sql&" where SOLICITUD.estado=1 "
sql = sql&" order by SOLICITUD.id_solicitud desc"

'RESPONSE.Write(sql)
'RESPONSE.End()

DATOS.Open sql,oConn
total_pages = 0
if( DATOS.RecordCount >0 )then 
	IF((DATOS.RecordCount MOD limite)>0)THEN
		total_pages = FIX(DATOS.RecordCount/limite )+1
	ELSE
		total_pages = (DATOS.RecordCount/limite)
	END IF	
END IF	



if (pagina > total_pages) then pagina=total_pages
inicio = (limite*pagina) - limite 
Response.Write("<?xml version='1.0' encoding='utf-8' ?>"&chr(13))
Response.Write("<rows>"&chr(13)) 
Response.Write("<page>"&pagina&"</page>"&chr(13))
Response.Write("<total>"&total_pages&"</total>"&chr(13))
Response.Write("<records>"&DATOS.RecordCount&"</records>"&chr(13))

'response.Write(sql)
'response.end()

fila=0
WHILE NOT DATOS.EOF
	if(fila>=inicio AND fila<(limite*pagina))then
		Response.Write("<row id="""&fila&""">"&chr(13))
		Response.Write("<cell><![CDATA["&DATOS("RUT")&"]]></cell>"&chr(13))
		Response.Write("<cell><![CDATA["&DATOS("empresa")&"]]></cell>"&chr(13))
		Response.Write("<cell><![CDATA["&DATOS("CODIGO")&"]]></cell>"&chr(13))
		Response.Write("<cell><![CDATA["&DATOS("NOMBRE_CURSO")&"]]></cell>"&chr(13))
		Response.Write("<cell><![CDATA["&DATOS("FECHA")&"]]></cell>"&chr(13))
		Response.Write("<cell><![CDATA[<span class=""ui-state-valid"" ><a href=""#"" title=""Ver Registro"" class=""ui-icon ui-icon-pencil"" name=""aContrato"" onclick=""update("&DATOS("id_solicitud")&")""></a></span>]]></cell>"&chr(13))
		Response.Write("</row>"&chr(13))
	end if
	fila=fila+1
	DATOS.MoveNext
WEND
Response.Write("</rows>") 
%>
