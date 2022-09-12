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
columna = "E.R_SOCIAL"
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

trab=""
if(Request("trabajador")<>"")then
trab=" and H.ID_TRABAJADOR='"&Request("trabajador")&"'"
end if

emp=""
if(Request("empresa")<>"")then
emp=" and H.ID_EMPRESA='"&Request("empresa")&"'"
end if

curso=""
if(Request("curso")<>"")then
curso=" and C.ID_MUTUAL='"&Request("curso")&"'"
end if

inicio=""
if(Request("inicio")<>"")then
	inicio=" and P.fecha_inicio_>=CONVERT(DATE, '"&Request("inicio")&"')"
end if

termino=""
if(Request("termino")<>"")then
	termino=" and P.FECHA_TERMINO<=CONVERT(DATE, '"&Request("termino")&"')"
end if

buscaIdMutual=" where H.EVALUACION='Aprobado' and H.ESTADO=0 "
if(request("u")="133")then
	buscaIdMutual=" where H.EVALUACION in ('Aprobado','Reprobado') and H.ESTADO in (0,2) "
end if

sql = "select E.ID_EMPRESA,E.RUT as RUTE,UPPER(E.R_SOCIAL) as R_SOCIAL,T.RUT,UPPER(T.NOMBRES) as NOMBRES,C.NOMBRE_CURSO, "
sql = sql&"CONVERT(VARCHAR(10),P.FECHA_TERMINO, 105) as FECHA_TERMINO,H.COD_AUTENFICACION,H.ID_PROGRAMA,H.RELATOR,"
sql = sql&"H.ID_TRABAJADOR,(CASE WHEN H.COD_AUTENFICACION is null then 'No Aplica' "
sql = sql&" WHEN H.COD_AUTENFICACION is not null then H.COD_AUTENFICACION END) as 'CODIGO', "
sql = sql&"(CASE WHEN H.ESTADO=2 then 'IPL' WHEN H.ESTADO=0 then 'ILB' END) as 'EINS',"
sql = sql&"(CASE WHEN H.EVALUACION='Aprobado' then 'A' WHEN H.EVALUACION='Reprobado' then 'R' END) as 'EsEval',C.ID_MUTUAL "
sql = sql&"  from HISTORICO_CURSOS H "
sql = sql&" inner join EMPRESAS E on E.ID_EMPRESA=H.ID_EMPRESA "
sql = sql&" inner join TRABAJADOR T on T.ID_TRABAJADOR=H.ID_TRABAJADOR "
sql = sql&" inner join PROGRAMA P on P.ID_PROGRAMA=H.ID_PROGRAMA "
sql = sql&" inner join CURRICULO C on C.ID_MUTUAL=P.ID_MUTUAL "
sql = sql&buscaIdMutual&trab&emp&curso&inicio&termino&" ORDER BY "&columna&" "&orden

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
		Response.Write("<cell><![CDATA["&replace(FormatNumber(mid(DATOS("RUTE"), 1,len(DATOS("RUTE"))-2),0)&mid(DATOS("RUTE"), len(DATOS("RUTE"))-1,len(DATOS("RUTE"))),",",".")&"]]></cell>"&chr(13))
		Response.Write("<cell><![CDATA["&DATOS("R_SOCIAL")&"]]></cell>"&chr(13))
		Response.Write("<cell><![CDATA["&replace(FormatNumber(mid(DATOS("RUT"), 1,len(DATOS("RUT"))-2),0)&mid(DATOS("RUT"), len(DATOS("RUT"))-1,len(DATOS("RUT"))),",",".")&"]]></cell>"&chr(13))
		Response.Write("<cell><![CDATA["&DATOS("NOMBRES")&"]]></cell>"&chr(13))
		Response.Write("<cell><![CDATA["&DATOS("FECHA_TERMINO")&"]]></cell>"&chr(13))
		Response.Write("<cell><![CDATA["&DATOS("CODIGO")&"]]></cell>"&chr(13))
		
		Color="#ff0000"
		if(DATOS("EsEval")="A")then
			Color="#66ff00"
		end if
		Response.Write("<cell><![CDATA[<span style=""color:#000000; position:relative;"">"&DATOS("EsEval")&"<span style=""color:"&Color&";top:-1px;left:-1px;position:absolute;"">"&DATOS("EsEval")&"</span></span>]]></cell>"&chr(13))
		
		Color="#ff0000"
		if(DATOS("EINS")="ILB")then
			Color="#66ff00"
		end if
		
		Response.Write("<cell><![CDATA[<span style=""color:#000000; position:relative;"">"&DATOS("EINS")&"<span style=""color:"&Color&";top:-1px;left:-1px;position:absolute;"">"&DATOS("EINS")&"</span></span>]]></cell>"&chr(13))
		
		if(DATOS("EINS")="ILB" and DATOS("EsEval")="A")then
			Response.Write("<cell><![CDATA[<span class=""ui-state-valid"" ><a href=""#"" title=""Ver Certificado"" class=""ui-icon ui-icon-document"" name=""aContrato"" onclick=""certificados("&DATOS("ID_PROGRAMA")&","&DATOS("ID_EMPRESA")&","&DATOS("ID_TRABAJADOR")&","&DATOS("RELATOR")&","&DATOS("ID_MUTUAL")&")""></a></span>]]></cell>"&chr(13))
		else
			Response.Write("<cell><![CDATA[<span class=""ui-state-valid"" ><a href=""#"" title=""No Aplica"" class=""ui-icon ui-icon-cancel"" name=""aContrato""></a></span>]]></cell>"&chr(13))		
		end if
		
		Response.Write("</row>"&chr(13))
	end if
	fila=fila+1
	DATOS.MoveNext
WEND
Response.Write("</rows>") 
%>
