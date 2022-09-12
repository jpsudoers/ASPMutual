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
columna = "CONVERT(date,PROGRAMA.FECHA_INICIO_, 105)"
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

	if(Session("userTipo")<>"")then
		   if(Session("userTipo")="1")then
		      filtraRel=" and (bloque_programacion.id_relator='"&Session("relUsuario")&"' or bloque_programacion.id_rel_seg='"&Session("relUsuario")&"')"
		   else
		      filtraRel=""
		   end if
	else
		filtraRel=" and bloque_programacion.id_relator=''"
	end if		   


sql ="select distinct HISTORICO_CURSOS.RELATOR,HISTORICO_CURSOS.SEDE,"
sql = sql&"dbo.MayMinTexto (INSTRUCTOR_RELATOR.NOMBRES+' '+INSTRUCTOR_RELATOR.A_PATERNO+' '+INSTRUCTOR_RELATOR.A_MATERNO) as instructor,"
sql = sql&"PROGRAMA.ID_PROGRAMA,CURRICULO.CODIGO,dbo.MayMinTexto(CURRICULO.NOMBRE_CURSO) as NOMBRE_CURSO,"
sql = sql&"CURRICULO.SENCE,CONVERT(VARCHAR(10),PROGRAMA.FECHA_INICIO_, 105) as FECHA_INICIO_,"
sql = sql&"(CASE WHEN bloque_programacion.id_rel_seg IS NOT NULL THEN ' / '+(select dbo.MayMinTexto (ri.NOMBRES+' '+ri.A_PATERNO+' '+ri.A_MATERNO) "
sql = sql&" from INSTRUCTOR_RELATOR ri where ri.ID_INSTRUCTOR=bloque_programacion.id_rel_seg) "
sql = sql&" WHEN bloque_programacion.id_rel_seg IS NULL THEN '' END) as rel_seg, "
sql = sql&" case when (select Datediff(""d"", Min(CONVERT(date,GETDATE(),105)), Max(CONVERT(date,pg.FECHA_TERMINO, 105))) "  
sql = sql&" from programa pg where pg.ID_PROGRAMA=PROGRAMA.ID_PROGRAMA)<=(-2) then '#ff0000' "  
sql = sql&" when (select Datediff(""d"", Min(CONVERT(date,GETDATE(),105)), Max(CONVERT(date,pg.FECHA_TERMINO, 105))) "  
sql = sql&" from programa pg where pg.ID_PROGRAMA=PROGRAMA.ID_PROGRAMA)=(-1) then '#ffff00' " 
sql = sql&" when (select Datediff(""d"", Min(CONVERT(date,GETDATE(),105)), Max(CONVERT(date,pg.FECHA_TERMINO, 105))) "  
sql = sql&" from programa pg where pg.ID_PROGRAMA=PROGRAMA.ID_PROGRAMA)>=0 then '#66ff00' end as color, "
sql = sql&" CONVERT(date,PROGRAMA.FECHA_INICIO_, 105),PROGRAMA.DIR_EJEC,"
sql = sql&" (case when GETDATE()>=DATEADD(hour, 17, PROGRAMA.FECHA_TERMINO) then '1' else '0' end) as estadoEval, "
sql = sql&" HISTORICO_CURSOS.ID_BLOQUE from PROGRAMA "
sql = sql&" inner join CURRICULO on CURRICULO.ID_MUTUAL=PROGRAMA.ID_MUTUAL "
sql = sql&" inner join HISTORICO_CURSOS on HISTORICO_CURSOS.ID_PROGRAMA=PROGRAMA.ID_PROGRAMA "
sql = sql&" inner join INSTRUCTOR_RELATOR on INSTRUCTOR_RELATOR.ID_INSTRUCTOR=HISTORICO_CURSOS.RELATOR "
sql = sql&" inner join bloque_programacion on bloque_programacion.id_bloque=HISTORICO_CURSOS.ID_BLOQUE "
sql = sql&"where HISTORICO_CURSOS.ESTADO=1 and bloque_programacion.estado=0 "&filtraRel
sql = sql&" ORDER BY "&columna&" "&orden
'sql = sql&" order by CONVERT(date,PROGRAMA.FECHA_INICIO_, 105) asc"

'RESPONSE.Write(sql)
'RESPONSE.End()

DATOS.Open sql,oConn
total_pages = 0
if( DATOS.RecordCount >0 )then 
	IF((DATOS.RecordCount MOD limite)>0)THEN
		total_pages = FIX(DATOS.RecordCount/limite )+1
	ELSE
		total_pages = (DATOS.RecordCount/limite)
	END IF	
END IF	



if (pagina > total_pages) then pagina=total_pages
inicio = (limite*pagina) - limite 
Response.Write("<?xml version='1.0' encoding='utf-8' ?>"&chr(13))
Response.Write("<rows>"&chr(13)) 
Response.Write("<page>"&pagina&"</page>"&chr(13))
Response.Write("<total>"&total_pages&"</total>"&chr(13))
Response.Write("<records>"&DATOS.RecordCount&"</records>"&chr(13))

'response.Write(sql)
'response.end()

fila=0
WHILE NOT DATOS.EOF
	if(fila>=inicio AND fila<(limite*pagina))then
		Response.Write("<row id="""&fila&""">"&chr(13))
		'Response.Write("<cell><![CDATA["&right("00000"&DATOS("ID_PROGRAMA"),5)&"]]></cell>"&chr(13))
		'Response.Write("<cell><![CDATA["&DATOS("FECHA_INICIO_")&"]]></cell>"&chr(13))
		
		Response.Write("<cell><![CDATA[<span style=""color:#000000; position:relative;"">"&DATOS("FECHA_INICIO_")&"<span style=""color:"&DATOS("color")&";top:-1px;left:-1px;position:absolute;"">"&DATOS("FECHA_INICIO_")&"</span></span>]]></cell>"&chr(13))					
		Response.Write("<cell><![CDATA["&DATOS("CODIGO")&"]]></cell>"&chr(13))
		'if(DATOS("SENCE")=0)then
		'Response.Write("<cell><![CDATA[Si]]></cell>"&chr(13))
		'else
		'Response.Write("<cell><![CDATA[No]]></cell>"&chr(13))
		'end if
		Response.Write("<cell><![CDATA["&DATOS("NOMBRE_CURSO")&"]]></cell>"&chr(13))
		Response.Write("<cell><![CDATA["&DATOS("instructor")&DATOS("rel_seg")&"]]></cell>"&chr(13))
		Response.Write("<cell><![CDATA["&DATOS("DIR_EJEC")&"]]></cell>"&chr(13))		
		Response.Write("<cell><![CDATA[<span class=""ui-state-valid"" ><a href=""#"" title=""Libro de Clases"" class=""ui-icon ui-icon-folder-collapsed"" name=""aContrato"" onclick=""update("&DATOS("ID_PROGRAMA")&","&DATOS("RELATOR")&","&DATOS("ID_BLOQUE")&")""></a></span>]]></cell>"&chr(13))
		
		if(DATOS("estadoEval")="1")then 
				Response.Write("<cell><![CDATA[<span class=""ui-state-valid"" ><a href=""#"" title=""Evaluar Curso"" class=""ui-icon ui-icon-check"" name=""aContrato"" onclick=""evaCierre("&DATOS("ID_PROGRAMA")&","&DATOS("RELATOR")&")""></a></span>]]></cell>"&chr(13))
		else
				Response.Write("<cell><![CDATA[<span class=""ui-state-valid"" ><a href=""#"" title=""Curso En Proceso"" class=""ui-icon ui-icon-cancel"" name=""aContrato""></a></span>]]></cell>"&chr(13))		
		end if
		
		Response.Write("<cell><![CDATA[<span class=""ui-state-valid""><a href=""#"" title=""Planilla Excel"" class=""ui-icon ui-icon-document"" name=""aContrato"" onclick=""excel("&DATOS("ID_PROGRAMA")&","&DATOS("RELATOR")&")""></a></span>]]></cell>"&chr(13))
		Response.Write("</row>"&chr(13))
	end if
	fila=fila+1
	DATOS.MoveNext
WEND
Response.Write("</rows>") 
%>
