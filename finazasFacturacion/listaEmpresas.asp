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
columna = "AUTORIZACION.FECHA_REV_CIERRE"
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


id_mutual=""
if(request("IDMutual")<>"")then
	id_mutual=" and CURRICULO.ID_MUTUAL='"&request("IDMutual")&"'"
end if


sql = "select EMPRESAS.RUT,UPPER (EMPRESAS.R_SOCIAL) as R_SOCIAL,EMPRESAS.ID_EMPRESA,AUTORIZACION.VALOR_OC,"
sql = sql&"EMPRESAS.ID_OTIC AS OTIC,AUTORIZACION.N_PARTICIPANTES,AUTORIZACION.CON_OTIC,"
sql = sql&"(select COUNT(*) from FACTURAS where FACTURAS.ID_AUTORIZACION=AUTORIZACION.ID_AUTORIZACION) as 'Facturado',"
sql = sql&"AUTORIZACION.DOCUMENTO_COMPROMISO,AUTORIZACION.ID_AUTORIZACION,AUTORIZACION.CON_FRANQUICIA,"
sql = sql&"bloque_programacion.id_programa,bloque_programacion.id_relator,"
sql = sql&"CONVERT(VARCHAR(10),AUTORIZACION.FECHA_REV_CIERRE, 105) AS FECHA_REV_CIERRE,"
sql = sql&" case (select Datediff(""d"", Min(CONVERT(date,AU.FECHA_REV_CIERRE,105)), Max(CONVERT(date,GETDATE(), 105))) " 
sql = sql&" from AUTORIZACION AU where AU.ID_AUTORIZACION=AUTORIZACION.ID_AUTORIZACION) " 
sql = sql&" when 0 then '#66ff00' " 
sql = sql&" when 1 then '#ffff00' " 
sql = sql&" when 2 then '#ff0000' " 
sql = sql&" else '#ff0000' end as color from AUTORIZACION "
sql = sql&" inner join EMPRESAS on EMPRESAS.ID_EMPRESA=AUTORIZACION.ID_EMPRESA "
sql = sql&" inner join bloque_programacion on bloque_programacion.id_bloque=AUTORIZACION.ID_BLOQUE "
sql = sql&" inner join PROGRAMA on PROGRAMA.ID_PROGRAMA=AUTORIZACION.ID_PROGRAMA "
sql = sql&" inner join CURRICULO on CURRICULO.ID_MUTUAL=PROGRAMA.ID_MUTUAL "
sql = sql&" where AUTORIZACION.ESTADO=0 AND AUTORIZACION.FACTURADO=1 "&id_mutual
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
		'Response.Write("<cell><![CDATA["&DATOS("FECHA_REV_CIERRE")&"]]></cell>"&chr(13))	
		Response.Write("<cell><![CDATA[<span style=""color:#000000; position:relative;"">"&DATOS("FECHA_REV_CIERRE")&"<span style=""color:"&DATOS("color")&";top:-1px;left:-1px;position:absolute;"">"&DATOS("FECHA_REV_CIERRE")&"</span></span>]]></cell>"&chr(13))	
		Response.Write("<cell><![CDATA["&replace(FormatNumber(mid(DATOS("RUT"), 1,len(DATOS("RUT"))-2),0)&mid(DATOS("RUT"), len(DATOS("RUT"))-1,len(DATOS("RUT"))),",",".")&"]]></cell>"&chr(13))
		Response.Write("<cell><![CDATA["&DATOS("R_SOCIAL")&"]]></cell>"&chr(13))
		Response.Write("<cell><![CDATA["&DATOS("N_PARTICIPANTES")&"]]></cell>"&chr(13))
		Response.Write("<cell><![CDATA["&"$"&replace(FormatNumber(DATOS("VALOR_OC"),0),",",".")&"]]></cell>"&chr(13))
		
		if(DATOS("CON_OTIC")="1")then
			Response.Write("<cell><![CDATA[<span class=""ui-state-valid"" ><a href=""#"" title=""con OTIC"" class=""ui-icon ui-icon-check"" name=""aContrato""></a></span>]]></cell>"&chr(13))
		else
			Response.Write("<cell><![CDATA[<span class=""ui-state-valid"" ><a href=""#"" title=""Sin OTIC"" class=""ui-icon ui-icon-close"" name=""aContrato""></a></span>]]></cell>"&chr(13))
		end if
		
		Response.Write("<cell><![CDATA[<span class=""ui-state-valid"" ><a href=""#"" title=""Revisar Registro"" class=""ui-icon ui-icon-pencil"" name=""aContrato"" onclick=""update("&DATOS("ID_AUTORIZACION")&")""></a></span>]]></cell>"&chr(13))
		
		Response.Write("<cell><![CDATA[<span class=""ui-state-valid"" ><a href=""#"" title=""Ver Documento"" class=""ui-icon ui-icon-document-b"" name=""aContrato"" onclick=""docOC('"&DATOS("DOCUMENTO_COMPROMISO")&"');""></a></span>]]></cell>"&chr(13))
		
		'if(DATOS("CON_FRANQUICIA")="1")then
			'Response.Write("<cell><![CDATA[<span class=""ui-state-valid"" ><a href=""#"" title=""Ver Certificado"" class=""ui-icon ui-icon-document"" name=""aContrato"" onclick=""docCert('"&"finazasFacturacion/CertSence.asp?IdAuto="&DATOS("ID_AUTORIZACION")&"')""></a></span>]]></cell>"&chr(13))
		'else
			'Response.Write("<cell><![CDATA[<span class=""ui-state-valid"" ><a href=""#"" title=""No Aplica"" class=""ui-icon ui-icon-circle-close"" name=""aContrato""></a></span>]]></cell>"&chr(13))
		'end if
		
		
		if(DATOS("Facturado")="0")then
		   if(DATOS("CON_OTIC")="1" AND DATOS("OTIC")<>"0")then
		    Response.Write("<cell><![CDATA[<span class=""ui-state-valid"" ><a href=""#"" title=""Imprimir Factura"" class=""ui-icon ui-icon-print"" name=""aContrato"" onclick=""detFacturaOtic("&DATOS("ID_AUTORIZACION")&")""></a></span>]]></cell>"&chr(13))
		   else
		    Response.Write("<cell><![CDATA[<span class=""ui-state-valid"" ><a href=""#"" title=""Imprimir Factura"" class=""ui-icon ui-icon-print"" name=""aContrato"" onclick=""detFactura("&DATOS("ID_AUTORIZACION")&")""></a></span>]]></cell>"&chr(13))
		   end if
		else
			Response.Write("<cell><![CDATA[<span class=""ui-state-valid"" ><a href=""#"" title=""Facturado"" class=""ui-icon ui-icon-circle-check"" name=""aContrato""></a></span>]]></cell>"&chr(13))
		end if

		Response.Write("</row>"&chr(13))
	end if
	fila=fila+1
	DATOS.MoveNext
WEND
Response.Write("</rows>") 
%>
