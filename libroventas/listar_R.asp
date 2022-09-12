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
columna = "F.FECHA_EMISION"
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
	prog=" and "
	if(Request("selMes")<>"0")then
		prog=prog&"Month(F.FECHA_EMISION)='"&Request("selMes")&"'"
	end if
	
	if(Request("selAno")<>"0")then
		if(Request("selMes")<>"0")then
			prog=prog&" and "
		end if
	
		prog=prog&" Year(F.FECHA_EMISION)='"&Request("selAno")&"'"
	end if

	'prog=" and A.ID_PROGRAMA in (select p2.ID_PROGRAMA from PROGRAMA p2 where "
	'if(Request("selMes")<>"0")then
		'prog=prog&"Month(p2.FECHA_INICIO_)='"&Request("selMes")&"'"
	'end if
	
	'if(Request("selAno")<>"0")then
		'if(Request("selMes")<>"0")then
			'prog=prog&" and "
		'end if
	
		'prog=prog&" Year(p2.FECHA_INICIO_)='"&Request("selAno")&"'"
	'end if
	
	'prog=prog&")"
end if

emp=""
if(Request("empresa")<>"")then
emp=" and F.ID_EMPRESA='"&Request("empresa")&"'"
end if

sql = "SELECT E.RUT as rut,dbo.MayMinTexto (E.R_SOCIAL) as empresa, "
sql = sql&"CONVERT(VARCHAR(10),P.FECHA_INICIO_, 105) as FECHA_INICIO,CONVERT(VARCHAR(10),F.FECHA_EMISION,105) as FECHA," 
sql = sql&"F.FECHA_EMISION,C.CODIGO, dbo.MayMinTexto (C.NOMBRE_CURSO) as nombre, A.N_PARTICIPANTES, "
sql = sql&"A.ID_AUTORIZACION,A.ID_PROGRAMA as ID_PROGRAMA,A.DOCUMENTO_COMPROMISO as doc, F.FACTURA,F.ID_FACTURA,"
sql = sql&"F.ESTADO,(CASE F.ESTADO WHEN 1 THEN 'Vigente' ELSE 'Anulada' end) as 'DESCEST' FROM FACTURAS F "
sql = sql&" inner join AUTORIZACION A on A.ID_AUTORIZACION=F.ID_AUTORIZACION "
sql = sql&" inner join EMPRESAS E on E.ID_EMPRESA=F.ID_EMPRESA "
sql = sql&" inner join PROGRAMA P on P.ID_PROGRAMA=A.ID_PROGRAMA "  
sql = sql&" inner join CURRICULO C on C.ID_MUTUAL=P.ID_MUTUAL " 
sql = sql&" inner join bloque_programacion BQ on BQ.id_bloque=A.ID_BLOQUE " 
sql = sql&" WHERE A.ESTADO=0 "&prog&emp&" ORDER BY "&columna&" "&orden


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
		Response.Write("<cell><![CDATA["&DATOS("FECHA")&"]]></cell>"&chr(13))
		Response.Write("<cell><![CDATA["&replace(FormatNumber(mid(DATOS("rut"), 1,len(DATOS("rut"))-2),0)&mid(DATOS("rut"), len(DATOS("rut"))-1,len(DATOS("rut"))),",",".")&"]]></cell>"&chr(13))
		Response.Write("<cell><![CDATA["&DATOS("empresa")&"]]></cell>"&chr(13))
		'Response.Write("<cell><![CDATA["&DATOS("CODIGO")&"]]></cell>"&chr(13))		
		'Response.Write("<cell><![CDATA["&DATOS("nombre")&"]]></cell>"&chr(13))
		Response.Write("<cell><![CDATA["&DATOS("FACTURA")&"]]></cell>"&chr(13))
		Response.Write("<cell><![CDATA["&DATOS("DESCEST")&"]]></cell>"&chr(13))		
        
		if(DATOS("ESTADO")="1")then
			Response.Write("<cell><![CDATA[<span class=""ui-state-valid"" ><a href=""#"" title=""Anular Factura"" class=""ui-icon ui-icon-arrowthickstop-1-s"" name=""aContrato"" onclick=""Anular("&DATOS("FACTURA")&","&DATOS("ID_AUTORIZACION")&","&DATOS("ID_FACTURA")&");""></a></span>]]></cell>"&chr(13))
			Response.Write("<cell><![CDATA[<span class=""ui-state-valid"" ><a href=""#"" title=""Ver Certificado"" class=""ui-icon ui-icon-document"" name=""aContrato"" onclick=""documento('"&"libroventas/CertSence.asp?IdAuto="&DATOS("ID_AUTORIZACION")&"&rutEmpresa="&DATOS("rut")&"&nomEmpresa="&replace(DATOS("empresa"),"&","y")&"&nFactura="&DATOS("FACTURA")&"')""></a></span>]]></cell>"&chr(13))	
		else
			Response.Write("<cell><![CDATA[<span class=""ui-state-valid"" ><a href=""#"" title=""Factura Anulada"" class=""ui-icon ui-icon-cancel"" name=""aContrato""></a></span>]]></cell>"&chr(13))
			Response.Write("<cell><![CDATA[<span class=""ui-state-valid"" ><a href=""#"" title=""Certificado No Disponible"" class=""ui-icon ui-icon-cancel"" name=""aContrato""></a></span>]]></cell>"&chr(13))
		end if	
		Response.Write("<cell><![CDATA[<span class=""ui-state-valid"" ><a href=""#"" title=""Ver Registro"" class=""ui-icon ui-icon-pencil"" name=""aContrato"" onclick=""update("&DATOS("ID_AUTORIZACION")&")""></a></span>]]></cell>"&chr(13))
		Response.Write("</row>"&chr(13))
	end if
	fila=fila+1
	DATOS.MoveNext
WEND
Response.Write("</rows>") 
%>