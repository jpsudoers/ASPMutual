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

orden = "DESC"
if(Request("sord") <> "")then orden = Request("sord")

columna = "CURRICULO.NOMBRE_CURSO"
if(Request("sidx") <> "")then columna = Request("sidx")

node=0
if(request("nodeid")<>"")then node=cInt(request("nodeid"))

%>

<%
Dim DATOS
Dim DATOS_CURSOS
Dim oConn
SET oConn = Server.CreateObject("ADODB.Connection")
'oConn.Open("Provider=SQLOLEDB; User ID=sa;Password=SCL.2013.2013;data source=.\SQLEXPRESS;Initial Catalog=dbmas")
oConn.Open(MM_cnn_STRING)
Set DATOS = Server.CreateObject("ADODB.RecordSet")
DATOS.CursorType=3

sql = "select PROGRAMA.ID_PROGRAMA,CURRICULO.CODIGO,(CASE WHEN CURRICULO.SENCE = 0 THEN 'Si' "
sql = sql&" WHEN CURRICULO.SENCE = 1 THEN 'No' END) as SENCE,dbo.MayMinTexto(CURRICULO.NOMBRE_CURSO) as NOMBRE_CURSO,"
sql = sql&"(PROGRAMA.CUPOS - (select COUNT(HISTORICO_CURSOS.ID_PROGRAMA) from HISTORICO_CURSOS "
sql = sql&"where HISTORICO_CURSOS.ID_PROGRAMA=PROGRAMA.ID_PROGRAMA)) as ins_vacantes, "
sql = sql&"CONVERT(VARCHAR(10),PROGRAMA.FECHA_CIERRE, 105) as FECHA_CIERRE, "
sql = sql&"CONVERT(VARCHAR(10),PROGRAMA.FECHA_INICIO_, 105) as FECHA_INICIO_, "
sql = sql&"(CASE WHEN PROGRAMA.DIR_EJEC <>  'NULL' THEN ' - '+PROGRAMA.DIR_EJEC else '' END) as direccion from PROGRAMA "
sql = sql&" inner join CURRICULO on CURRICULO.ID_MUTUAL=PROGRAMA.ID_MUTUAL "
sql = sql&" where PROGRAMA.ESTADO=1 AND PROGRAMA.VIGENCIA=1 "
sql = sql&" and PROGRAMA.FECHA_CIERRE>=CONVERT(date,GETDATE(), 105) "
sql = sql&" and (PROGRAMA.CUPOS - (select COUNT(HISTORICO_CURSOS.ID_PROGRAMA) from HISTORICO_CURSOS "
sql = sql&" where HISTORICO_CURSOS.ID_PROGRAMA=PROGRAMA.ID_PROGRAMA)) > 0  AND TIPO=2  " 
sql = sql&" ORDER BY "&columna&" "&orden

DATOS.Open sql,oConn
'total_pages = 0
'if( DATOS.RecordCount >0 )then 
	'IF((DATOS.RecordCount MOD limite)>0)THEN
		'total_pages = FIX(DATOS.RecordCount/limite )+1
	'ELSE
		'total_pages = (DATOS.RecordCount/limite)
	'END IF	
'END IF	

Set DATOS_CURSOS = Server.CreateObject("ADODB.RecordSet")
DATOS_CURSOS.CursorType=3

sql2 = "select PROGRAMA.ID_PROGRAMA,PROGRAMA.ID_EMPRESA,CURRICULO.CODIGO,(CASE WHEN CURRICULO.SENCE = 0 THEN 'Si' "
sql2 = sql2&" WHEN CURRICULO.SENCE = 1 THEN 'No' END) as SENCE,dbo.MayMinTexto(CURRICULO.NOMBRE_CURSO) as NOMBRE_CURSO,"
sql2 = sql2&"(PROGRAMA.CUPOS - (select COUNT(HISTORICO_CURSOS.ID_PROGRAMA) from HISTORICO_CURSOS "
sql2 = sql2&"where HISTORICO_CURSOS.ID_PROGRAMA=PROGRAMA.ID_PROGRAMA)) as ins_vacantes, "
sql2 = sql2&"CONVERT(VARCHAR(10),PROGRAMA.FECHA_CIERRE, 105) as FECHA_CIERRE, "
sql2 = sql2&"CONVERT(VARCHAR(10),PROGRAMA.FECHA_INICIO_, 105) as FECHA_INICIO_, "
sql2 = sql2&"(CASE WHEN PROGRAMA.ID_EMPRESA is Null THEN '' WHEN PROGRAMA.ID_EMPRESA is not Null "
sql2 = sql2&" THEN ' - '+(select em.R_SOCIAL from EMPRESAS em where em.ID_EMPRESA=PROGRAMA.ID_EMPRESA) END) as r_social, "
sql2 = sql2&"(CASE WHEN PROGRAMA.DIR_EJEC <>  'NULL' THEN ' ('+PROGRAMA.DIR_EJEC+')' else '' END) as direccion from PROGRAMA "
sql2 = sql2&" inner join CURRICULO on CURRICULO.ID_MUTUAL=PROGRAMA.ID_MUTUAL "
sql2 = sql2&" where PROGRAMA.ESTADO=1 AND PROGRAMA.VIGENCIA=1 "
sql2 = sql2&" and PROGRAMA.FECHA_CIERRE>=CONVERT(date,GETDATE(), 105) "
sql2 = sql2&" and (PROGRAMA.CUPOS - (select COUNT(HISTORICO_CURSOS.ID_PROGRAMA) from HISTORICO_CURSOS "
sql2 = sql2&" where HISTORICO_CURSOS.ID_PROGRAMA=PROGRAMA.ID_PROGRAMA)) > 0 AND TIPO in (1,3) "
sql2 = sql2&" ORDER BY "&columna&" "&orden

DATOS_CURSOS.Open sql2,oConn

total_filas = 0
total_pages = 0

total_filas = DATOS.RecordCount + DATOS_CURSOS.RecordCount
if( total_filas >0 )then 
	IF((total_filas MOD limite)>0)THEN
		total_pages = FIX(total_filas/limite )+1
	ELSE
		total_pages = (total_filas/limite)
	END IF	
END IF	


if (pagina > total_pages) then pagina=total_pages
inicio = (limite*pagina) - limite 
Response.Write("<?xml version='1.0' encoding='utf-8' ?>"&chr(13))
Response.Write("<rows>"&chr(13)) 
Response.Write("<page>"&pagina&"</page>"&chr(13))
Response.Write("<total>"&total_pages&"</total>"&chr(13))
Response.Write("<records>"&total_filas&"</records>"&chr(13))

fila=0
WHILE NOT DATOS.EOF
		if(fila>=inicio AND fila<(limite*pagina))then
			Response.Write("<row id="""&fila&""">"&chr(13))
			Response.Write("<cell><![CDATA["&right("000000000"&DATOS("ID_PROGRAMA"),5)&"]]></cell>"&chr(13))
			Response.Write("<cell><![CDATA["&DATOS("CODIGO")&"]]></cell>"&chr(13))
			'Response.Write("<cell><![CDATA["&DATOS("SENCE")&"]]></cell>"&chr(13))
			Response.Write("<cell><![CDATA["&DATOS("NOMBRE_CURSO")&DATOS("direccion")&"]]></cell>"&chr(13))
			Response.Write("<cell><![CDATA["&DATOS("ins_vacantes")&"]]></cell>"&chr(13))
			Response.Write("<cell><![CDATA["&DATOS("FECHA_INICIO_")&"]]></cell>"&chr(13))
			Response.Write("<cell><![CDATA[<span class=""ui-state-valid"" ><a href=""#"" title=""Modificar Registro"" class=""ui-icon ui-icon-pencil"" name=""aContrato"" onclick=""update("&DATOS("ID_PROGRAMA")&","&DATOS("ins_vacantes")&",2,0)""></a></span>]]></cell>"&chr(13))
			Response.Write("</row>"&chr(13)) 
		end if
	fila=fila+1
	DATOS.MoveNext
WEND

WHILE NOT DATOS_CURSOS.EOF
		if(fila>=inicio AND fila<(limite*pagina))then
			Response.Write("<row id="""&fila&""">"&chr(13))
			Response.Write("<cell><![CDATA[<b>"&right("000000000"&DATOS_CURSOS("ID_PROGRAMA"),5)&"</b>]]></cell>"&chr(13))
			Response.Write("<cell><![CDATA[<b>"&DATOS_CURSOS("CODIGO")&"</b>]]></cell>"&chr(13))
			'Response.Write("<cell><![CDATA[<b>"&DATOS_CURSOS("SENCE")&"</b>]]></cell>"&chr(13))
			Response.Write("<cell><![CDATA[<b>"&DATOS_CURSOS("NOMBRE_CURSO")&DATOS_CURSOS("r_social")&DATOS_CURSOS("direccion")&"</b>]]></cell>"&chr(13))
			Response.Write("<cell><![CDATA[<b>"&DATOS_CURSOS("ins_vacantes")&"</b>]]></cell>"&chr(13))
			Response.Write("<cell><![CDATA[<b>"&DATOS_CURSOS("FECHA_INICIO_")&"</b>]]></cell>"&chr(13))
			Response.Write("<cell><![CDATA[<span class=""ui-state-valid"" ><a href=""#"" title=""Modificar Registro"" class=""ui-icon ui-icon-pencil"" name=""aContrato"" onclick=""update("&DATOS_CURSOS("ID_PROGRAMA")&","&DATOS_CURSOS("ins_vacantes")&",1,"&DATOS_CURSOS("ID_EMPRESA")&")""></a></span>]]></cell>"&chr(13))
			Response.Write("</row>"&chr(13)) 
		end if
	fila=fila+1
	DATOS_CURSOS.MoveNext
WEND

Response.Write("</rows>") 
%>
