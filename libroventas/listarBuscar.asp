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

bdoc=""
if(Request("txDoc")<>"")then
	bdoc=" and F.factura='"&Request("txDoc")&"'"
end if

txfcd=""
if(Request("fcd")<>"")then
	txfcd=" and P.FECHA_INICIO_>=CONVERT(DATE,'"&Request("fcd")&"')"
end if

txfch=""
if(Request("fch")<>"")then
	txfch=" and P.FECHA_INICIO_<=CONVERT(DATE,'"&Request("fch")&"')"
end if

txfed=""
if(Request("fed")<>"")then
	txfed=" and F.FECHA_EMISION>=CONVERT(DATE,'"&Request("fed")&"')"
end if

txfeh=""
if(Request("feh")<>"")then
	txfeh=" and F.FECHA_EMISION<=CONVERT(DATE,'"&Request("feh")&"')"
end if

txe=""
if(Request("e")<>"")then
	txe=" and F.ID_EMPRESA='"&Request("e")&"'"
end if

txc=""
if(Request("c")<>"0")then
	txc=" and C.ID_MUTUAL='"&Request("c")&"'"
end if

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
sql = sql&" WHERE A.ESTADO in (0,1) "&bdoc&txfcd&txfch&txfed&txfeh&txe&txc&" ORDER BY "&columna&" "&orden


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
		'Response.Write("<cell><![CDATA["&DATOS("ID_FACTURA")&"]]></cell>"&chr(13))
		Response.Write("<cell><![CDATA["&DATOS("FECHA")&"]]></cell>"&chr(13))
		Response.Write("<cell><![CDATA["&replace(FormatNumber(mid(DATOS("rut"), 1,len(DATOS("rut"))-2),0)&mid(DATOS("rut"), len(DATOS("rut"))-1,len(DATOS("rut"))),",",".")&"]]></cell>"&chr(13))
		Response.Write("<cell><![CDATA["&DATOS("empresa")&"]]></cell>"&chr(13))
		'Response.Write("<cell><![CDATA["&DATOS("CODIGO")&"]]></cell>"&chr(13))		
		'Response.Write("<cell><![CDATA["&DATOS("nombre")&"]]></cell>"&chr(13))
		Response.Write("<cell><![CDATA["&DATOS("FACTURA")&"]]></cell>"&chr(13))
		Response.Write("<cell><![CDATA["&DATOS("DESCEST")&"]]></cell>"&chr(13))		
        
		if(DATOS("ESTADO")="1" or DATOS("ESTADO")="2")then
			'Response.Write("<cell><![CDATA[<span class=""ui-state-valid"" ><a href=""#"" title=""Anular o Refacturar"" class=""ui-icon ui-icon-arrowthickstop-1-s"" name=""aContrato"" onclick=""Anular("&DATOS("FACTURA")&","&DATOS("ID_AUTORIZACION")&","&DATOS("ID_FACTURA")&");""></a></span>]]></cell>"&chr(13))
			Response.Write("<cell><![CDATA[<span class=""ui-state-valid"" ><a href=""#"" title=""Ver Certificado"" class=""ui-icon ui-icon-document"" name=""aContrato"" onclick=""documento('"&"libroventas/CertSence.asp?IdAuto="&DATOS("ID_AUTORIZACION")&"&rutEmpresa="&DATOS("RUT_EMPRESA")&"&nomEmpresa="&replace(DATOS("NOMBRE_EMPRESA"),"&","y")&"&nFactura="&DATOS("FACTURA")&"&rutOtic="&DATOS("rut_otic")&"&nomOtic="&DATOS("nombre_otic")&"')""></a></span>]]></cell>"&chr(13))
			'if(DATOS("tipo_emp")="2")then 
				'Response.Write("<cell><![CDATA[<span class=""ui-state-valid""><a href=""#"" title=""Ver Ticket"" class=""ui-icon ui-icon-document-b"" name=""aContrato"" onclick=""tickets('"&DATOS("ID_AUTORIZACION")&"','"&DATOS("N_TICKETS")&"','"&DATOS("FACTURA")&"')""></a></span>]]></cell>"&chr(13))
			'else
				'Response.Write("<cell><![CDATA[<span class=""ui-state-valid"" ><a href=""#"" title=""Ticket No Aplica"" class=""ui-icon ui-icon-cancel"" name=""aContrato""></a></span>]]></cell>"&chr(13))		
			'end if
		else
			'Response.Write("<cell><![CDATA[<span class=""ui-state-valid"" ><a href=""#"" title=""Factura Anulada"" class=""ui-icon ui-icon-cancel"" name=""aContrato""></a></span>]]></cell>"&chr(13))
			Response.Write("<cell><![CDATA[<span class=""ui-state-valid"" ><a href=""#"" title=""Certificado No Disponible"" class=""ui-icon ui-icon-cancel"" name=""aContrato""></a></span>]]></cell>"&chr(13))
			'Response.Write("<cell><![CDATA[<span class=""ui-state-valid"" ><a href=""#"" title=""Ticket No Disponible"" class=""ui-icon ui-icon-cancel"" name=""aContrato""></a></span>]]></cell>"&chr(13))			
		end if	
		Response.Write("<cell><![CDATA[<span class=""ui-state-valid"" ><a href=""#"" title=""Ver Registro"" class=""ui-icon ui-icon-pencil"" name=""aContrato"" onclick=""update("&DATOS("ID_AUTORIZACION")&")""></a></span>]]></cell>"&chr(13))
		Response.Write("</row>"&chr(13))
	end if
	fila=fila+1
	DATOS.MoveNext
WEND
Response.Write("</rows>") 
%>