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
columna = "CONVERT(date,P.FECHA_INICIO_)"
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

sql ="select distinct H.ID_AUTORIZACION,E.RUT as rut,dbo.MayMinTexto (E.R_SOCIAL) as empresa, " 
sql = sql&"CONVERT(VARCHAR(10),P.FECHA_INICIO_, 105) as FECHA_INICIO,dbo.MayMinTexto (C.NOMBRE_CURSO) as nombre,H.ID_BLOQUE," 
sql = sql&"CONVERT(date,P.FECHA_INICIO_),A.DOCUMENTO_COMPROMISO as doc from HISTORICO_CURSOS H "
sql = sql&" inner join AUTORIZACION A on A.ID_AUTORIZACION=H.ID_AUTORIZACION "
sql = sql&" inner join EMPRESAS E on E.ID_EMPRESA=H.ID_EMPRESA "
sql = sql&" inner join PROGRAMA P on P.ID_PROGRAMA=H.ID_PROGRAMA "  
sql = sql&" inner join CURRICULO C on C.ID_MUTUAL=P.ID_MUTUAL "
sql = sql&" where H.ESTADO=2 AND (CASE " 
sql = sql&"  WHEN A.CON_OTIC=1 AND A.CON_FRANQUICIA=1 AND A.TIPO_DOC=4 "
sql = sql&" AND (A.N_REG_SENCE IS NULL OR A.N_REG_SENCE IS NOT NULL) then '0' " 
sql = sql&"  WHEN A.CON_OTIC=0 AND A.CON_FRANQUICIA=1 AND A.TIPO_DOC=4 AND A.N_REG_SENCE IS NOT NULL then '0' " 
sql = sql&"  WHEN A.CON_OTIC=0 AND A.CON_FRANQUICIA=0 AND A.TIPO_DOC=4 then '0' " 
sql = sql&"  WHEN A.CON_OTIC=0 AND A.CON_FRANQUICIA=1 AND A.TIPO_DOC<>4 then '1' " 
sql = sql&"  WHEN A.CON_OTIC=0 AND A.CON_FRANQUICIA=0 AND A.TIPO_DOC<>4 then '1' "  
sql = sql&"  WHEN A.CON_OTIC=1 AND A.CON_FRANQUICIA=1 AND A.TIPO_DOC<>4 AND A.N_REG_SENCE IS NOT NULL then '1' END)=1 "
sql = sql&" ORDER BY "&columna&" "&orden
'sql = sql&" ORDER BY CONVERT(date,P.FECHA_INICIO_),H.ID_BLOQUE ASC "

'ORDER BY "&columna&" "&orden

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
		Response.Write("<cell><![CDATA["&DATOS("nombre")&"]]></cell>"&chr(13))
		Response.Write("<cell><![CDATA[<span class=""ui-state-valid"" ><a href=""#"" title=""Ver Documento"" class=""ui-icon ui-icon-document"" name=""aContrato"" onclick=""documento('"&DATOS("doc")&"');""></a></span>]]></cell>"&chr(13))
		
		
		'if(DATOS("TotCartasCompromiso")="0")then
		Response.Write("<cell><![CDATA[<span class=""ui-state-valid"" ><a href=""#"" title=""Revisión y Cierre"" class=""ui-icon ui-icon-unlocked"" name=""aContrato"" onclick=""update("&DATOS("ID_AUTORIZACION")&")""></a></span>]]></cell>"&chr(13))
		'else
		'Response.Write("<cell><![CDATA[<span class=""ui-state-valid"" ><a href=""#"" title=""Bloqueado"" class=""ui-icon ui-icon-cancel"" name=""aContrato""></a></span>]]></cell>"&chr(13))
		'end if
		
		Response.Write("</row>"&chr(13))
	end if
	fila=fila+1
	DATOS.MoveNext
WEND
Response.Write("</rows>") 
%>
