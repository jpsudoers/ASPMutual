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
columna = "preinscripciones.id_programacion"
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

sql = "select preinscripciones.id_preinscripcion as preinscripcion,preinscripciones.id_programacion as programa, "
sql = sql&"dbo.MayMinTexto(EMPRESAS.R_SOCIAL) as empresa,EMPRESAS.RUT as rut,dbo.MayMinTexto(CURRICULO.NOMBRE_CURSO) as curso, "
sql = sql&"CONVERT(VARCHAR(10),preinscripciones.fecha_preinscripcion, 105) as fecha_preinscripcion, "
sql = sql&"preinscripciones.participantes, "
sql = sql&" case (select Datediff(""d"", Min(CONVERT(date,ps.fecha_preinscripcion,105)), Max(CONVERT(date,GETDATE(), 105))) " 
sql = sql&" from preinscripciones ps where ps.id_preinscripcion=preinscripciones.id_preinscripcion) " 
sql = sql&" when 0 then '#66ff00' " 
sql = sql&" when 1 then '#ffff00' " 
sql = sql&" when 2 then '#ff0000' " 
sql = sql&" else '#ff0000' end as color from preinscripciones "
sql = sql&" inner join PROGRAMA on PROGRAMA.ID_PROGRAMA=preinscripciones.id_programacion "
sql = sql&" inner join EMPRESAS on EMPRESAS.ID_EMPRESA=preinscripciones.id_empresa "
sql = sql&" inner join CURRICULO on CURRICULO.ID_MUTUAL=PROGRAMA.ID_MUTUAL "
sql = sql&" where preinscripciones.estado=1 "
sql = sql&" ORDER BY "&columna&" "&orden
'sql = sql&" order by preinscripciones.id_preinscripcion desc "


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

session.lcid=1033
fila=0
WHILE NOT DATOS.EOF
	if(fila>=inicio AND fila<(limite*pagina))then
		Response.Write("<row id="""&fila&""">"&chr(13))
		Response.Write("<cell><![CDATA[<span style=""color:#000000; position:relative;"">"&DATOS("fecha_preinscripcion")&"<span style=""color:"&DATOS("color")&";top:-1px;left:-1px;position:absolute;"">"&DATOS("fecha_preinscripcion")&"</span></span>]]></cell>"&chr(13))		
		Response.Write("<cell><![CDATA["&right("000000000"&DATOS("programa"),5)&"]]></cell>"&chr(13))
		Response.Write("<cell><![CDATA["&DATOS("rut")&"]]></cell>"&chr(13))
		Response.Write("<cell><![CDATA["&DATOS("empresa")&"]]></cell>"&chr(13))
		Response.Write("<cell><![CDATA["&DATOS("curso")&"]]></cell>"&chr(13))
		'Response.Write("<cell><![CDATA["&DATOS("fecha_preinscripcion")&"]]></cell>"&chr(13))
	
		Response.Write("<cell><![CDATA["&DATOS("participantes")&"]]></cell>"&chr(13))
		Response.Write("<cell><![CDATA[<span class=""ui-state-valid"" ><a href=""#"" title=""Modificar Registro"" class=""ui-icon ui-icon-pencil"" name=""aContrato"" onclick=""update("&DATOS("preinscripcion")&")""></a></span>]]></cell>"&chr(13))
        Response.Write("<cell><![CDATA[<span class=""ui-state-valid"" ><a href=""#"" title=""Eliminar Registro"" class=""ui-icon ui-icon-trash"" name=""aContrato"" onclick=""eliminar("&DATOS("preinscripcion")&")""></a></span>]]></cell>"&chr(13))
       	Response.Write("<cell><![CDATA[<span class=""ui-state-valid"" ><a href=""#"" title=""Rechazar Registro"" class=""ui-icon ui-icon-arrowthickstop-1-s"" name=""aContrato"" onclick=""rechazarIns("&DATOS("preinscripcion")&")""></a></span>]]></cell>"&chr(13))				
		Response.Write("</row>"&chr(13))
	end if
	fila=fila+1
	DATOS.MoveNext
WEND
Response.Write("</rows>") 
%>
