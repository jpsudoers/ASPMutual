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
oConn.Open("Provider=SQLOLEDB; User ID=sa;Password=SCL.2013.2013;data source=.\SQLEXPRESS;Initial Catalog=dbmas")
Set DATOS = Server.CreateObject("ADODB.RecordSet")
DATOS.CursorType=3

sql = "select PROGRAMA.ID_PROGRAMA,CURRICULO.CODIGO,CURRICULO.SENCE as SENCE,CURRICULO.NOMBRE_CURSO,PROGRAMA.CUPOS,"
sql = sql&"CONVERT(VARCHAR(10),PROGRAMA.FECHA_INICIO_, 105) as FECHA_INICIO_,"
sql = sql&"(select COUNT(*) from HISTORICO_CURSOS where HISTORICO_CURSOS.ID_PROGRAMA=PROGRAMA.ID_PROGRAMA) as inscritos "
sql = sql&" from PROGRAMA "
sql = sql&" inner join CURRICULO on CURRICULO.ID_MUTUAL=PROGRAMA.ID_MUTUAL"
sql = sql&" where PROGRAMA.ESTADO=1 AND PROGRAMA.VIGENCIA=1"
sql = sql&" ORDER BY "&columna&" "&orden

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

'response.Write(sql)
'response.end()

fila=0
WHILE NOT DATOS.EOF
	if(fila>=inicio AND fila<(limite*pagina))then
		Response.Write("<row id="""&fila&""">"&chr(13))
		Response.Write("<cell><![CDATA["&right("000000000"&DATOS("ID_PROGRAMA"),5)&"]]></cell>"&chr(13))
		Response.Write("<cell><![CDATA["&DATOS("CODIGO")&"]]></cell>"&chr(13))
		'if(DATOS("SENCE")=0)then
		'Response.Write("<cell><![CDATA[Si]]></cell>"&chr(13))
		'else
		'Response.Write("<cell><![CDATA[No]]></cell>"&chr(13))
		'end if
		Response.Write("<cell><![CDATA["&DATOS("NOMBRE_CURSO")&"]]></cell>"&chr(13))
		Response.Write("<cell><![CDATA["&DATOS("CUPOS")&"]]></cell>"&chr(13))
		Response.Write("<cell><![CDATA["&DATOS("inscritos")&"]]></cell>"&chr(13))
		Response.Write("<cell><![CDATA["&DATOS("FECHA_INICIO_")&"]]></cell>"&chr(13))
		Response.Write("<cell><![CDATA[<span class=""ui-state-valid"" ><a href=""#"" title=""Modificar Registro"" class=""ui-icon ui-icon-pencil"" name=""aContrato"" onclick=""update("&DATOS("ID_PROGRAMA")&")""></a></span>]]></cell>"&chr(13))
Response.Write("<cell><![CDATA[<span class=""ui-state-valid"" ><a href=""#"" title=""Eliminar Registro"" class=""ui-icon ui-icon-trash"" name=""aContrato"" onclick=""eliminar("&DATOS("ID_PROGRAMA")&")""></a></span>]]></cell>"&chr(13))
		Response.Write("</row>"&chr(13))
	end if
	fila=fila+1
	DATOS.MoveNext
WEND
Response.Write("</rows>") 
%>
