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
columna = "PROGRAMA.FECHA_INICIO_"
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

prog=""
if(Request("selMes")<>"0" or Request("selAno")<>"0")then
	prog=" and AUTORIZACION.ID_PROGRAMA in (select p.ID_PROGRAMA from PROGRAMA p where "
	if(Request("selMes")<>"0")then
		prog=prog&"Month(p.FECHA_INICIO_)='"&Request("selMes")&"'"
	end if
	
	if(Request("selAno")<>"0")then
		if(Request("selMes")<>"0")then
			prog=prog&" and "
		end if
	
		prog=prog&" Year(p.FECHA_INICIO_)='"&Request("selAno")&"'"
	end if
	
	prog=prog&")"
end if

emp=" and AUTORIZACION.ID_EMPRESA=0"
if(Request("empresa")<>"0")then
emp=" and AUTORIZACION.ID_EMPRESA='"&Request("empresa")&"'"
end if

sql = "select EMPRESAS.RUT as rut,dbo.MayMinTexto (EMPRESAS.R_SOCIAL) as empresa, "
sql = sql&"CONVERT(VARCHAR(10),PROGRAMA.FECHA_INICIO_, 105) as FECHA_INICIO, "
sql = sql&"CURRICULO.CODIGO, dbo.MayMinTexto (CURRICULO.NOMBRE_CURSO) as nombre, AUTORIZACION.N_PARTICIPANTES,"
sql = sql&" AUTORIZACION.ID_AUTORIZACION,AUTORIZACION.ID_PROGRAMA as ID_PROGRAMA,AUTORIZACION.DOCUMENTO_COMPROMISO as doc "
sql = sql&" from AUTORIZACION "
sql = sql&" inner join PROGRAMA on PROGRAMA.ID_PROGRAMA=AUTORIZACION.ID_PROGRAMA " 
sql = sql&" inner join EMPRESAS on EMPRESAS.ID_EMPRESA=AUTORIZACION.ID_EMPRESA " 
sql = sql&" inner join CURRICULO on CURRICULO.ID_MUTUAL=PROGRAMA.ID_MUTUAL "
sql = sql&" inner join bloque_programacion on bloque_programacion.id_bloque=AUTORIZACION.ID_BLOQUE "
sql = sql&" WHERE AUTORIZACION.ESTADO=0 "&prog&emp&" ORDER BY "&columna&" "&orden

'response.Write(sql)
'response.End()

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
		Response.Write("<cell><![CDATA["&DATOS("FECHA_INICIO")&"]]></cell>"&chr(13))
		Response.Write("<cell><![CDATA["&DATOS("rut")&"]]></cell>"&chr(13))
		Response.Write("<cell><![CDATA["&DATOS("empresa")&"]]></cell>"&chr(13))
		Response.Write("<cell><![CDATA["&DATOS("CODIGO")&"]]></cell>"&chr(13))		
		Response.Write("<cell><![CDATA["&DATOS("nombre")&"]]></cell>"&chr(13))
		Response.Write("<cell><![CDATA["&DATOS("N_PARTICIPANTES")&"]]></cell>"&chr(13))
		Response.Write("<cell><![CDATA[<span class=""ui-state-valid"" ><a href=""#"" title=""Ver Documento"" class=""ui-icon ui-icon-document"" name=""aContrato"" onclick=""documento('"&DATOS("doc")&"');""></a></span>]]></cell>"&chr(13))
		Response.Write("<cell><![CDATA[<span class=""ui-state-valid"" ><a href=""#"" title=""Modificar Registro"" class=""ui-icon ui-icon-pencil"" name=""aContrato"" onclick=""update("&DATOS("ID_AUTORIZACION")&")""></a></span>]]></cell>"&chr(13))
		Response.Write("<cell><![CDATA[<span class=""ui-state-valid"" ><a href=""#"" title=""Detalle Facturas"" class=""ui-icon ui-icon-document-b"" name=""aContrato"" onclick=""detFactura("&DATOS("ID_AUTORIZACION")&")""></a></span>]]></cell>"&chr(13))		
		Response.Write("</row>"&chr(13))
	end if
	fila=fila+1
	DATOS.MoveNext
WEND
Response.Write("</rows>") 
%>