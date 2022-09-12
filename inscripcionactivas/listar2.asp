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
columna = "PROGRAMA.ID_PROGRAMA"
if(Request("sidx") <> "")then columna = Request("sidx")

node=0
if(request("nodeid")<>"")then node=cInt(request("nodeid"))

%>

<%
Dim DATOS
Dim DATOSTRAB
Dim oConn
SET oConn = Server.CreateObject("ADODB.Connection")
'oConn.Open("Provider=SQLOLEDB; User ID=sa;Password=SCL.2013.2013;data source=.\SQLEXPRESS;Initial Catalog=dbmas")
oConn.Open(MM_cnn_STRING)
Set DATOS = Server.CreateObject("ADODB.RecordSet")
DATOS.CursorType=3

Set DATOSTRAB = Server.CreateObject("ADODB.RecordSet")
DATOSTRAB.CursorType=3

sql = "select TRABAJADOR.RUT,TRABAJADOR.NOMBRES,TRABAJADOR.CARGO_EMPRESA,TRABAJADOR.ESCOLARIDAD,"
sql = sql&"HISTORICO_CURSOS.ID_HISTORICO_CURSO,HISTORICO_CURSOS.ID_TRABAJADOR,HISTORICO_CURSOS.ID_PROGRAMA,"
sql = sql&"HISTORICO_CURSOS.TRABUP,HISTORICO_CURSOS.TRABIDUP "
sql = sql&" from HISTORICO_CURSOS inner join TRABAJADOR on TRABAJADOR.ID_TRABAJADOR=HISTORICO_CURSOS.ID_TRABAJADOR "
sql = sql&" where HISTORICO_CURSOS.ID_EMPRESA="&Request("empresa")&" and HISTORICO_CURSOS.ID_PROGRAMA="&Request("programa")
sql = sql&" and HISTORICO_CURSOS.ID_AUTORIZACION="&Request("autorizacion")
sql = sql&" and HISTORICO_CURSOS.TRABDEL<>"&Request("tabPart")
sql = sql&" order by TRABAJADOR.NOMBRES ASC"

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
inicio = (limite*pagina) - limite 
Response.Write("<?xml version='1.0' encoding='utf-8' ?>"&chr(13))
Response.Write("<rows>"&chr(13)) 
Response.Write("<page>"&pagina&"</page>"&chr(13))
Response.Write("<total>"&total_pages&"</total>"&chr(13))
Response.Write("<records>"&DATOS.RecordCount&"</records>"&chr(13))

fila=0
WHILE NOT DATOS.EOF
	if(fila>=inicio AND fila<(limite*pagina))then
		Response.Write("<row id="""&fila&""">"&chr(13))
		
		if(cstr(DATOS("TRABUP"))=cstr(Request("tabPart")))then
		query = "select TRABAJADOR.RUT,TRABAJADOR.NOMBRES,TRABAJADOR.CARGO_EMPRESA,TRABAJADOR.ESCOLARIDAD,"
		query = query&"TRABAJADOR.ID_TRABAJADOR from TRABAJADOR "
		query = query&" where TRABAJADOR.ID_TRABAJADOR='"&DATOS("TRABIDUP")&"'"

		DATOSTRAB.Open query,oConn
		
		Response.Write("<cell><![CDATA["&DATOSTRAB("RUT")&"]]></cell>"&chr(13))
		Response.Write("<cell><![CDATA["&DATOSTRAB("NOMBRES")&"]]></cell>"&chr(13))
		Response.Write("<cell><![CDATA["&DATOSTRAB("CARGO_EMPRESA")&"]]></cell>"&chr(13))
		
		if(DATOSTRAB("ESCOLARIDAD")=0)then
		Response.Write("<cell><![CDATA[Sin Escolaridad]]></cell>"&chr(13))
		end if
		
		if(DATOSTRAB("ESCOLARIDAD")=1)then
		Response.Write("<cell><![CDATA[Básica Incompleta]]></cell>"&chr(13))
		end if
		
		if(DATOSTRAB("ESCOLARIDAD")=2)then
		Response.Write("<cell><![CDATA[Básica Completa]]></cell>"&chr(13))
		end if

		if(DATOSTRAB("ESCOLARIDAD")=3)then
		Response.Write("<cell><![CDATA[Media Incompleta]]></cell>"&chr(13))
		end if

		if(DATOSTRAB("ESCOLARIDAD")=4)then
		Response.Write("<cell><![CDATA[Media Completa]]></cell>"&chr(13))
		end if
		
		if(DATOSTRAB("ESCOLARIDAD")=5)then
		Response.Write("<cell><![CDATA[Superior Técnica Incompleta]]></cell>"&chr(13))
		end if

		if(DATOSTRAB("ESCOLARIDAD")=6)then
		Response.Write("<cell><![CDATA[Superior Técnica Profesional Completa]]></cell>"&chr(13))
		end if

		if(DATOSTRAB("ESCOLARIDAD")=7)then
		Response.Write("<cell><![CDATA[Universitaria Incompleta]]></cell>"&chr(13))
		end if
		
		if(DATOSTRAB("ESCOLARIDAD")=8)then
		Response.Write("<cell><![CDATA[Universitaria Completa]]></cell>"&chr(13))
		end if
		
	Response.Write("<cell><![CDATA[<span class=""ui-state-valid"" ><a href=""#"" title=""Cambiar Participante"" class=""ui-icon ui-icon-transferthick-e-w"" name=""aContrato"" onclick=""reempTrab('"&DATOS("ID_HISTORICO_CURSO")&"');""></a></span>]]></cell>"&chr(13))
		
		Response.Write("<cell><![CDATA[<span class=""ui-state-valid"" ><a href=""#"" title=""Modificar Registro"" class=""ui-icon ui-icon-pencil"" name=""aContrato"" onclick=""upTrabajador('"&DATOSTRAB("ID_TRABAJADOR")&"');""></a></span>]]></cell>"&chr(13))
		
		Response.Write("<cell><![CDATA[<span class=""ui-state-valid"" ><a href=""#"" title=""Eliminar Registro"" class=""ui-icon ui-icon-trash"" name=""aContrato"" onclick=""delTrab('"&DATOS("ID_HISTORICO_CURSO")&"');""></a></span>]]></cell>"&chr(13))
		
		DATOSTRAB.Close
		
		else
		
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
		
	Response.Write("<cell><![CDATA[<span class=""ui-state-valid"" ><a href=""#"" title=""Cambiar Participante"" class=""ui-icon ui-icon-transferthick-e-w"" name=""aContrato"" onclick=""reempTrab('"&DATOS("ID_HISTORICO_CURSO")&"');""></a></span>]]></cell>"&chr(13))
		
		Response.Write("<cell><![CDATA[<span class=""ui-state-valid"" ><a href=""#"" title=""Modificar Registro"" class=""ui-icon ui-icon-pencil"" name=""aContrato"" onclick=""upTrabajador('"&DATOS("ID_TRABAJADOR")&"');""></a></span>]]></cell>"&chr(13))
		
		Response.Write("<cell><![CDATA[<span class=""ui-state-valid"" ><a href=""#"" title=""Eliminar Registro"" class=""ui-icon ui-icon-trash"" name=""aContrato"" onclick=""delTrab('"&DATOS("ID_HISTORICO_CURSO")&"');""></a></span>]]></cell>"&chr(13))

		end if
		
		Response.Write("</row>"&chr(13))
		end if
		fila=fila+1
		DATOS.MoveNext
WEND
Response.Write("</rows>") 
%>
