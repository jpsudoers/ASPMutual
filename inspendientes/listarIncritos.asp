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
columna = "1"
if(Request("sidx") <> "")then columna = Request("sidx")

node=0
if(request("nodeid")<>"")then node=cInt(request("nodeid"))
%>

<%
Dim DATOS
Dim oConn
SET oConn = Server.CreateObject("ADODB.Connection")
oConn.Open(MM_cnn_STRING)
Set DATOS = Server.CreateObject("ADODB.RecordSet")
DATOS.CursorType=3

vprog=Request("prog")
'vrelator=Request("relator")
'vsede=Request("sede")

Dim vrelator
vrelator = "0"
If (Request("relator") <> "") Then 
  vrelator = Request("relator")
End If

Dim vsede
vsede = "0"
If (Request("sede") <> "") Then 
  vsede = Request("sede")
End If

sql = "select TRABAJADOR.RUT,TRABAJADOR.NOMBRES,TRABAJADOR.CARGO_EMPRESA,TRABAJADOR.ESCOLARIDAD from HISTORICO_CURSOS "
sql = sql&"  inner join TRABAJADOR on TRABAJADOR.ID_TRABAJADOR=HISTORICO_CURSOS.ID_TRABAJADOR "
sql = sql&"  where HISTORICO_CURSOS.ID_PROGRAMA="&vprog&" and HISTORICO_CURSOS.RELATOR="&vrelator
sql = sql&"  and HISTORICO_CURSOS.SEDE="&vsede
sql = sql&" ORDER BY "&columna&" "&orden

DATOS.Open sql,oConn
total_pages = 0
if( DATOS.RecordCount >0 )then 
	IF((DATOS.RecordCount MOD limite)>0)THEN
		total_pages = FIX(DATOS.RecordCount/limite )+1
	ELSE
		total_pages = (DATOS.RecordCount/limite)
	END IF	
ELSE
		total_pages = 1
END IF	
if (pagina > total_pages) then pagina=total_pages
inicio = (limite*pagina) - limite ' do not put $limit*($page - 1) 
Response.Write("<?xml version='1.0' encoding='utf-8' ?>"&chr(13))
Response.Write("<rows>"&chr(13)) 
Response.Write("<page>"&pagina&"</page>"&chr(13))
Response.Write("<total>"&total_pages&"</total>"&chr(13))
Response.Write("<records>"&DATOS.RecordCount&"</records>"&chr(13))

fila=0
WHILE NOT DATOS.EOF
	if(fila>=inicio AND fila<(limite*pagina))then
		Response.Write("<row id="""&fila&""">"&chr(13))
		Response.Write("<cell><![CDATA["&DATOS("RUT")&"]]></cell>"&chr(13))
		Response.Write("<cell><![CDATA["&DATOS("NOMBRES")&"]]></cell>"&chr(13))
		Response.Write("<cell><![CDATA["&DATOS("CARGO_EMPRESA")&"]]></cell>"&chr(13))
		
		if(DATOS("ESCOLARIDAD")=0)then
		Response.Write("<cell><![CDATA[Sin Escolaridad]]></cell>"&chr(13))
		end if
		
		if(DATOS("ESCOLARIDAD")=1)then
		Response.Write("<cell><![CDATA[Básica Incompleta]]></cell>"&chr(13))
		end if
		
		if(DATOS("ESCOLARIDAD")=2)then
		Response.Write("<cell><![CDATA[Básica Completa]]></cell>"&chr(13))
		end if

		if(DATOS("ESCOLARIDAD")=3)then
		Response.Write("<cell><![CDATA[Media Incompleta]]></cell>"&chr(13))
		end if

		if(DATOS("ESCOLARIDAD")=4)then
		Response.Write("<cell><![CDATA[Media Completa]]></cell>"&chr(13))
		end if
		
		if(DATOS("ESCOLARIDAD")=5)then
		Response.Write("<cell><![CDATA[Superior Técnica Incompleta]]></cell>"&chr(13))
		end if

		if(DATOS("ESCOLARIDAD")=6)then
		Response.Write("<cell><![CDATA[Superior Técnica Profesional Completa]]></cell>"&chr(13))
		end if

		if(DATOS("ESCOLARIDAD")=7)then
		Response.Write("<cell><![CDATA[Universitaria Incompleta]]></cell>"&chr(13))
		end if
		
		if(DATOS("ESCOLARIDAD")=8)then
		Response.Write("<cell><![CDATA[Universitaria Completa]]></cell>"&chr(13))
		end if

		Response.Write("</row>"&chr(13))
		end if
	fila=fila+1
	DATOS.MoveNext
WEND
Response.Write("</rows>") 
%>

