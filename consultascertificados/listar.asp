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
columna = "T.NOMBRES"
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
if(Request("trabajador")<>"0" and Request("trabajador")<>"")then
trab=" and H.ID_TRABAJADOR='"&Request("trabajador")&"'"
end if

emp=""
if(Request("emp")<>"0" and Request("emp")<>"")then
emp=" and H.ID_EMPRESA='"&Request("emp")&"'"
end if

curso=" and P.ID_MUTUAL='"&Request("curso")&"'"

sql = "select E.ID_EMPRESA,dbo.MayMinTexto(E.R_SOCIAL) as R_SOCIAL,T.RUT,dbo.MayMinTexto(T.NOMBRES) as NOMBRES,C.NOMBRE_CURSO, "
sql = sql&"CONVERT(VARCHAR(10),P.FECHA_TERMINO, 105) as FECHA_TERMINO,H.ID_PROGRAMA,H.RELATOR,"
sql = sql&"H.ID_TRABAJADOR,(CASE WHEN H.COD_AUTENFICACION is null then 'No Aplica' "
sql = sql&" WHEN H.COD_AUTENFICACION is not null then H.COD_AUTENFICACION END) as 'CODIGO',"
sql = sql&"dbo.MayMinTexto(C.NOMBRE_CURSO) as NOMBRE_CURSO, "
sql = sql&"CONVERT(VARCHAR(10),DATEADD(MONTH, (CASE " 
sql = sql&" WHEN C.VIGENCIA =  1 THEN 6 "
sql = sql&" WHEN C.VIGENCIA =  2 THEN 12 "
sql = sql&" WHEN C.VIGENCIA =  3 THEN 18 "
sql = sql&" WHEN C.VIGENCIA =  4 THEN 24 "
sql = sql&" WHEN C.VIGENCIA =  5 THEN 48 "
sql = sql&" END), CONVERT(date,DATEADD(DAY, 1, CONVERT(date,P.FECHA_TERMINO, 105)), 105)), 105) as vigencia,"
sql = sql&"(CASE WHEN CONVERT(date,GETDATE(), 105)<=DATEADD(MONTH, (CASE " 
sql = sql&"		 WHEN C.VIGENCIA =  1 THEN 6 "
sql = sql&" 	 WHEN C.VIGENCIA =  2 THEN 12 "
sql = sql&"		 WHEN C.VIGENCIA =  3 THEN 18 "
sql = sql&"		 WHEN C.VIGENCIA =  4 THEN 24 "
sql = sql&"		 WHEN C.VIGENCIA =  5 THEN 48 "
sql = sql&"	     END), CONVERT(date,DATEADD(DAY, 1, CONVERT(date,P.FECHA_TERMINO, 105)), 105)) THEN '#000' "
sql = sql&"WHEN CONVERT(date,GETDATE(), 105)>DATEADD(MONTH, (CASE " 
sql = sql&"		 WHEN C.VIGENCIA =  1 THEN 6 "
sql = sql&"		 WHEN C.VIGENCIA =  2 THEN 12 "
sql = sql&"		 WHEN C.VIGENCIA =  3 THEN 18 "
sql = sql&"		 WHEN C.VIGENCIA =  4 THEN 24 "
sql = sql&"		 WHEN C.VIGENCIA =  5 THEN 48 "
sql = sql&"		 END), CONVERT(date,DATEADD(DAY, 1, CONVERT(date,P.FECHA_TERMINO, 105)), 105)) THEN '#ff0000' "
sql = sql&"END) as color,C.ID_MUTUAL, h.id_historico_curso "  
sql = sql&"  from HISTORICO_CURSOS H "
sql = sql&" inner join EMPRESAS E on E.ID_EMPRESA=H.ID_EMPRESA "
sql = sql&" inner join TRABAJADOR T on T.ID_TRABAJADOR=H.ID_TRABAJADOR "
sql = sql&" inner join PROGRAMA P on P.ID_PROGRAMA=H.ID_PROGRAMA "
sql = sql&" inner join CURRICULO C on C.ID_MUTUAL=P.ID_MUTUAL "
sql = sql&" where H.EVALUACION<>'Reprobado' and H.ESTADO=0 "&trab&emp&curso&" ORDER BY "&columna&" "&orden

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
		Response.Write("<cell>"&DATOS("id_historico_curso")&"</cell>"&chr(13))		
		Response.Write("<cell><![CDATA[<span style=""color:"&DATOS("color")&""">"&replace(FormatNumber(mid(DATOS("RUT"), 1,len(DATOS("RUT"))-2),0)&mid(DATOS("RUT"), len(DATOS("RUT"))-1,len(DATOS("RUT"))),",",".")&"</span>]]></cell>"&chr(13))
		Response.Write("<cell><![CDATA[<span style=""color:"&DATOS("color")&""">"&DATOS("NOMBRES")&"</span>]]></cell>"&chr(13))
	   Response.Write("<cell><![CDATA[<span style=""color:"&DATOS("color")&""">"&DATOS("R_SOCIAL")&"</span>]]></cell>"&chr(13))
	  Response.Write("<cell><![CDATA[<span style=""color:"&DATOS("color")&""">"&DATOS("FECHA_TERMINO")&"</span>]]></cell>"&chr(13))		
		'Response.Write("<cell><![CDATA[<span style=""color:"&DATOS("color")&""">"&DATOS("vigencia")&"</span>]]></cell>"&chr(13))			
		
		if(DATOS("CODIGO")<>"No Aplica")then
		Response.Write("<cell><![CDATA[<span class=""ui-state-valid"" ><a href=""#"" title=""Ver Certificado"" class=""ui-icon ui-icon-document"" name=""aContrato"" onclick=""certificados("&DATOS("ID_PROGRAMA")&","&DATOS("ID_EMPRESA")&","&DATOS("ID_TRABAJADOR")&","&DATOS("RELATOR")&","&DATOS("ID_MUTUAL")&")""></a></span>]]></cell>"&chr(13))
		else
		Response.Write("<cell><![CDATA[<span class=""ui-state-valid"" ><a href=""#"" title=""Certificado No Disponible"" class=""ui-icon ui-icon-cancel"" name=""aContrato""></a></span>]]></cell>"&chr(13))
		end if
		Response.Write("</row>"&chr(13))
	end if
	fila=fila+1
	DATOS.MoveNext
WEND
Response.Write("</rows>") 
%>

