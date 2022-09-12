<!--#include file="cnn_string.asp"-->
<%

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
'oConn.Open("Provider=SQLOLEDB; User ID=sa;Password=SCL.2013.2013;data source=.\SQLEXPRESS;Initial Catalog=dbmas")
oConn.Open(MM_cnn_STRING)
Set DATOS = Server.CreateObject("ADODB.RecordSet")
DATOS.CursorType=3



EstaL=""
if(Request("estado")="1")then
	EstaL=EstaL&" and (select COUNT(*) from verifica_inscripcion VI where VI.ID_AUTORIZACION=A.ID_AUTORIZACION AND VI.ESTADO=1)=0 "
elseif(Request("estado")="2")then
	EstaL=EstaL&" and (select COUNT(*) from verifica_inscripcion VI where VI.ID_AUTORIZACION=A.ID_AUTORIZACION AND VI.ESTADO=1)>0 "
else
  	EstaL=""
end if

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
sql = sql&"  and H.ID_PROGRAMA in (select p2.ID_PROGRAMA from PROGRAMA p2 "
sql = sql&" where CONVERT(date,p2.FECHA_INICIO_)>CONVERT(date,'25-04-2011')) "
sql = sql&" and ((isnull(e.CON_REFERENCIA,0)=1 and isnull(a.ID_TIPO_REFERENCIA,0)>0 and isnull(a.N_REFERENCIA,'')<>'' "
sql = sql&" and a.TIPO_DOC=0) OR (isnull(e.CON_REFERENCIA,0)=1 and a.TIPO_DOC<>0) or isnull(e.CON_REFERENCIA,0)=0) "

if(Request("tipoVenta") <> "") THEN
sql = sql&" and A.ID_TIPO_VENTA = "&Request("tipoVenta")
end if

if(Request("condicion") <> "null" and Request("condicion") <> "")  THEN
sql = sql&" and A.TIPO_DOC = "&Request("condicion")
end if

if(Request("empB") <> "null" and Request("empB") <> "")  THEN
sql = sql&" and A.ID_EMPRESA = "&Request("empB")
end if

sql = sql&" union all "
sql = sql&" select a.ID_AUTORIZACION,E.RUT as rut,dbo.MayMinTexto (E.R_SOCIAL) as empresa, " 
sql = sql&"  CONVERT(VARCHAR(10),P.FECHA_INICIO_, 105) as FECHA_INICIO,dbo.MayMinTexto (C.NOMBRE_CURSO) as nombre,a.ID_BLOQUE,   " 
sql = sql&"  CONVERT(date,P.FECHA_INICIO_),A.DOCUMENTO_COMPROMISO from autorizacion a  " 
sql = sql&"   inner join EMPRESAS E on E.ID_EMPRESA=a.ID_EMPRESA   " 
sql = sql&"   inner join PROGRAMA P on P.ID_PROGRAMA=a.ID_PROGRAMA     " 
sql = sql&"   inner join CURRICULO C on C.ID_MUTUAL=P.ID_MUTUAL   " 
sql = sql&"   where a.ESTADO=1 and (a.SOLO_CERTIFICADOS=1 or a.SOLO_CERTIFICADOS is null) " 
sql = sql&"   and (select count(*) from HISTORICO_CURSOS hc  " 
sql = sql&"   where hc.ID_AUTORIZACION=a.ID_AUTORIZACION and hc.ESTADO=0)=a.N_PARTICIPANTES " 
sql = sql&"    AND (CASE    " 
sql = sql&"    WHEN A.CON_OTIC=1 AND A.CON_FRANQUICIA=1 AND A.TIPO_DOC=4   " 
sql = sql&"   AND (A.N_REG_SENCE IS NULL OR A.N_REG_SENCE IS NOT NULL) then '0'    " 
sql = sql&"    WHEN A.CON_OTIC=0 AND A.CON_FRANQUICIA=1 AND A.TIPO_DOC=4 AND A.N_REG_SENCE IS NOT NULL then '0'   "  
sql = sql&"    WHEN A.CON_OTIC=0 AND A.CON_FRANQUICIA=0 AND A.TIPO_DOC=4 then '0'    " 
sql = sql&"    WHEN A.CON_OTIC=0 AND A.CON_FRANQUICIA=1 AND A.TIPO_DOC<>4 then '1'    " 
sql = sql&"    WHEN A.CON_OTIC=0 AND A.CON_FRANQUICIA=0 AND A.TIPO_DOC<>4 then '1'     " 
sql = sql&"    WHEN A.CON_OTIC=1 AND A.CON_FRANQUICIA=1 AND A.TIPO_DOC<>4 AND A.N_REG_SENCE IS NOT NULL then '1' END)=1 "   
sql = sql&"    and a.ID_PROGRAMA in (select p2.ID_PROGRAMA from PROGRAMA p2   " 
sql = sql&"   where CONVERT(date,p2.FECHA_INICIO_)>CONVERT(date,'25-04-2011'))   " 
sql = sql&"   and ((isnull(e.CON_REFERENCIA,0)=1 and isnull(a.ID_TIPO_REFERENCIA,0)>0 and isnull(a.N_REFERENCIA,'')<>''  "  
sql = sql&"   and a.TIPO_DOC=0) OR (isnull(e.CON_REFERENCIA,0)=1 and a.TIPO_DOC<>0) or isnull(e.CON_REFERENCIA,0)=0) "

if(Request("tipoVenta") <> "") THEN
sql = sql&" and A.ID_TIPO_VENTA = "&Request("tipoVenta")
end if

if(Request("condicion") <> "null" and Request("condicion") <> "")  THEN
sql = sql&" and A.TIPO_DOC = "&Request("condicion")
end if

if(Request("empB") <> "null" and Request("empB") <> "")  THEN
sql = sql&" and A.ID_EMPRESA = "&Request("empB")
end if

sql = sql&EstaL&" ORDER BY "&columna&" "&orden
'sql = sql&" ORDER BY CONVERT(date,P.FECHA_INICIO_),H.ID_BLOQUE ASC "

'ORDER BY "&columna&" "&orden
Response.Write(sql)
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
		Response.Write("<cell><![CDATA[<span class=""ui-state-valid"" ><a href=""#"" title=""Revisi󮠹 Cierre"" class=""ui-icon ui-icon-unlocked"" name=""aContrato"" onclick=""update("&DATOS("ID_AUTORIZACION")&")""></a></span>]]></cell>"&chr(13))
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
