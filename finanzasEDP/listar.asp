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
columna = "FECHAs"
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


queryMes = ""
queryAno = ""
queryEmpresa = ""
selfmes = request("selMes")
selAno = request("selAno")
rutEmpresa = request("empresa")
tipoDoc = request("txDoc")
estado = request("SelEst")


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

est=""
if(Request("SelEst")<>"")then
	est=" and F.ESTADO='"&Request("SelEst")&"'"
end if

bdoc=""
if(Request("txDoc")<>"")then
	bdoc=" and F.N_DOCUMENTO LIKE '%"&Request("txDoc")&"%'"
end if


 'sql = "select a.ID_FACTURA, a.FECHA_EMISION as FECHA "
 'sql = sql&", b.RUT as rut, b.R_SOCIAL as empresa"
 'sql = sql&", a.FACTURA as FACTURA"
 'sql = sql&", 'Pendiente de Pago' as Estado "
 'sql = sql&" from FACTURAS a"
 'sql = sql&" inner join EMPRESAS b on a.ID_EMPRESA = b.ID_EMPRESA"
 'sql = sql&" inner join EMPRESA_CONDICION_COMERCIAL c on b.ID_EMPRESA = c.ID_EMPRESA"
 'sql = sql&" where c.ID_CONDICION_COMERCIAL in (0,5) and a.FECHA_PAGO is null and DESCRIPCION_ESTADO is null "
 'sql = sql&queryEmpresa&queryAno&queryMes
 sql = "SELECT E.RUT as rut,dbo.MayMinTexto (E.R_SOCIAL) as empresa, "
sql = sql&"CONVERT(VARCHAR(10),P.FECHA_INICIO_, 105) as FECHA_INICIO,CONVERT(VARCHAR(10),F.FECHA_EMISION,105) as FECHA," 
sql = sql&"F.FECHA_EMISION,C.CODIGO, dbo.MayMinTexto (C.NOMBRE_CURSO) as nombre, A.N_PARTICIPANTES, "
sql = sql&"A.ID_AUTORIZACION,A.ID_PROGRAMA as ID_PROGRAMA,A.DOCUMENTO_COMPROMISO as doc, F.FACTURA,F.ID_FACTURA,"
sql = sql&"F.ESTADO,(CASE F.ESTADO WHEN 1 THEN 'Vigente' WHEN 0 THEN 'Anulada' WHEN 2 THEN 'Vig/Ref' end) as 'DESCEST',"
sql = sql&" EI.RUT AS RUT_EMPRESA,EI.R_SOCIAL AS NOMBRE_EMPRESA,"
sql = sql&"(CASE WHEN A.CON_FRANQUICIA=1 and A.CON_OTIC=1 and A.ID_OTIC<>0 THEN " 
sql = sql&"(select eotic.RUT from EMPRESAS eotic where eotic.ID_EMPRESA=A.ID_OTIC) ELSE 'No Aplica' end) as rut_otic, "
sql = sql&"(CASE WHEN A.CON_FRANQUICIA=1 and A.CON_OTIC=1 and A.ID_OTIC<>0 THEN " 
sql = sql&"(select dbo.MayMinTexto (eotic.R_SOCIAL) from EMPRESAS eotic where eotic.ID_EMPRESA=A.ID_OTIC) "
sql = sql&" ELSE 'No Aplica' end) as nombre_otic,E.TIPO as tipo_emp,N_TICKETS FROM FACTURAS F "
sql = sql&" inner join AUTORIZACION A on A.ID_AUTORIZACION=F.ID_AUTORIZACION "
sql = sql&" inner join EMPRESAS EI ON EI.ID_EMPRESA=A.ID_EMPRESA "
sql = sql&" inner join EMPRESAS E on E.ID_EMPRESA=F.ID_EMPRESA "
sql = sql&" inner join PROGRAMA P on P.ID_PROGRAMA=A.ID_PROGRAMA "  
sql = sql&" inner join CURRICULO C on C.ID_MUTUAL=P.ID_MUTUAL " 
sql = sql&" inner join bloque_programacion BQ on BQ.id_bloque=A.ID_BLOQUE " 
sql = sql&" inner join  EMPRESA_CONDICION_COMERCIAL cc on E.ID_EMPRESA = cc.ID_EMPRESA"
sql = sql&" WHERE A.ESTADO in (0,1) and  cc.ID_CONDICION_COMERCIAL in (0,5) and F.FECHA_PAGO is null and F.DESCRIPCION_ESTADO is null "&prog&emp&est&bdoc&" ORDER BY "&columna&" "&orden
 


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

'fila=0
'WHILE NOT DATOS.EOF
'	if(fila>=inicio AND fila<(limite*pagina))then
'		Response.Write("<row id="""&fila&""">"&chr(13))
'		Response.Write("<cell><![CDATA["&DATOS("ID_FACTURA")&"]]></cell>"&chr(13))
'		Response.Write("<cell><![CDATA["&DATOS("FECHA")&"]]></cell>"&chr(13))
'		Response.Write("<cell><![CDATA["&replace(FormatNumber(mid(DATOS("rut"), 1,len(DATOS("rut"))-2),0)&mid(DATOS("rut"), len(DATOS("rut"))-1,len(DATOS("rut"))),",",".")&"]]></cell>"&chr(13))
'		Response.Write("<cell><![CDATA["&DATOS("empresa")&"]]></cell>"&chr(13))
'		Response.Write("<cell><![CDATA["&DATOS("FACTURA")&"]]></cell>"&chr(13))
'		Response.Write("<cell><![CDATA["&DATOS("Estado")&"]]></cell>"&chr(13))
'		Response.Write("</row>"&chr(13))
'	end if
'	fila=fila+1
'	DATOS.MoveNext
'WEND
'Response.Write("</rows>") 
fila=0
WHILE NOT DATOS.EOF
	if(fila>=inicio AND fila<(limite*pagina))then
		Response.Write("<row id="""&fila&""">"&chr(13))
		Response.Write("<cell><![CDATA["&DATOS("ID_FACTURA")&"]]></cell>"&chr(13))
		Response.Write("<cell><![CDATA["&DATOS("FECHA")&"]]></cell>"&chr(13))
		Response.Write("<cell><![CDATA["&replace(FormatNumber(mid(DATOS("rut"), 1,len(DATOS("rut"))-2),0)&mid(DATOS("rut"), len(DATOS("rut"))-1,len(DATOS("rut"))),",",".")&"]]></cell>"&chr(13))
		Response.Write("<cell><![CDATA["&DATOS("empresa")&"]]></cell>"&chr(13))
		'Response.Write("<cell><![CDATA["&DATOS("CODIGO")&"]]></cell>"&chr(13))		
		'Response.Write("<cell><![CDATA["&DATOS("nombre")&"]]></cell>"&chr(13))
		Response.Write("<cell><![CDATA["&DATOS("FACTURA")&"]]></cell>"&chr(13))
		Response.Write("<cell><![CDATA["&DATOS("DESCEST")&"]]></cell>"&chr(13))		
        
		if(DATOS("ESTADO")="1" or DATOS("ESTADO")="2")then
			Response.Write("<cell><![CDATA[<span class=""ui-state-valid"" ><a href=""#"" title=""Anular o Refacturar"" class=""ui-icon ui-icon-arrowthickstop-1-s"" name=""aContrato"" onclick=""Anular("&DATOS("FACTURA")&","&DATOS("ID_AUTORIZACION")&","&DATOS("ID_FACTURA")&");""></a></span>]]></cell>"&chr(13))
			Response.Write("<cell><![CDATA[<span class=""ui-state-valid"" ><a href=""#"" title=""Ver Certificado"" class=""ui-icon ui-icon-document"" name=""aContrato"" onclick=""documento('"&"libroventas/CertSence.asp?IdAuto="&DATOS("ID_AUTORIZACION")&"&rutEmpresa="&DATOS("RUT_EMPRESA")&"&nomEmpresa="&replace(DATOS("NOMBRE_EMPRESA"),"&","y")&"&nFactura="&DATOS("FACTURA")&"&rutOtic="&DATOS("rut_otic")&"&nomOtic="&DATOS("nombre_otic")&"')""></a></span>]]></cell>"&chr(13))
			'if(DATOS("tipo_emp")="2")then 
				'Response.Write("<cell><![CDATA[<span class=""ui-state-valid""><a href=""#"" title=""Ver Ticket"" class=""ui-icon ui-icon-document-b"" name=""aContrato"" onclick=""tickets('"&DATOS("ID_AUTORIZACION")&"','"&DATOS("N_TICKETS")&"','"&DATOS("FACTURA")&"')""></a></span>]]></cell>"&chr(13))
			'else
				'Response.Write("<cell><![CDATA[<span class=""ui-state-valid"" ><a href=""#"" title=""Ticket No Aplica"" class=""ui-icon ui-icon-cancel"" name=""aContrato""></a></span>]]></cell>"&chr(13))		
			'end if
		else
			Response.Write("<cell><![CDATA[<span class=""ui-state-valid"" ><a href=""#"" title=""Factura Anulada"" class=""ui-icon ui-icon-cancel"" name=""aContrato""></a></span>]]></cell>"&chr(13))
			Response.Write("<cell><![CDATA[<span class=""ui-state-valid"" ><a href=""#"" title=""Certificado No Disponible"" class=""ui-icon ui-icon-cancel"" name=""aContrato""></a></span>]]></cell>"&chr(13))
			'Response.Write("<cell><![CDATA[<span class=""ui-state-valid"" ><a href=""#"" title=""Ticket No Disponible"" class=""ui-icon ui-icon-cancel"" name=""aContrato""></a></span>]]></cell>"&chr(13))			
		end if	
		Response.Write("<cell><![CDATA[<span class=""ui-state-valid"" ><a href=""#"" title=""Ver Registro"" class=""ui-icon ui-icon-pencil"" name=""aContrato"" onclick=""update("&DATOS("ID_AUTORIZACION")&")""></a></span>]]></cell>"&chr(13))
		
		if(DATOS("FACTURA")<>"")then
				Response.Write("<cell><![CDATA[<span class=""ui-state-valid"" ><a href=""#"" title=""No Disponible"" class=""ui-icon ui-icon-cancel"" name=""aContrato""></a></span>]]></cell>"&chr(13))	
				Response.Write("<cell><![CDATA[<span class=""ui-state-valid"" ><a href=""#"" title=""No Disponible"" class=""ui-icon ui-icon-cancel"" name=""aContrato""></a></span>]]></cell>"&chr(13))					
		else
				if(Request("u_nv")="1")then
						Response.Write("<cell><![CDATA[<span class=""ui-state-valid"" ><a href=""#"" title=""Eliminar Nota Venta"" class=""ui-icon ui-icon-circle-close"" name=""aContrato"" onclick=""DelNotaVenta("&DATOS("ID_FACTURA")&","&DATOS("ID_AUTORIZACION")&")""></a></span>]]></cell>"&chr(13))	
						Response.Write("<cell><![CDATA[<span class=""ui-state-valid"" ><a href=""#"" title=""Asigna Número de Factura"" class=""ui-icon ui-icon-newwin"" name=""aContrato"" onclick=""asgNFactura("&DATOS("ID_FACTURA")&")""></a></span>]]></cell>"&chr(13))								
				else
						Response.Write("<cell><![CDATA[<span class=""ui-state-valid"" ><a href=""#"" title=""No Disponible"" class=""ui-icon ui-icon-circle-close"" name=""aContrato""></a></span>]]></cell>"&chr(13))	
						Response.Write("<cell><![CDATA[<span class=""ui-state-valid"" ><a href=""#"" title=""No Disponible"" class=""ui-icon ui-icon-newwin"" name=""aContrato""></a></span>]]></cell>"&chr(13))							
				end if			
		end if 	
		
		Response.Write("</row>"&chr(13))
	end if
	fila=fila+1
	DATOS.MoveNext
WEND
Response.Write("</rows>") 
%>