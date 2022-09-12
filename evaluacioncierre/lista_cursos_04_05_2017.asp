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

vrelator=0
if(request("r")<>"0")then vrelator=request("r") end if

vtipouser=0
if(request("tp")<>"0")then vtipouser=request("tp") end if

vusr=0
if(request("usr")<>"0")then vusr=request("usr") end if
%>

<%
Dim DATOS
Dim oConn
SET oConn = Server.CreateObject("ADODB.Connection")
'oConn.Open("Provider=SQLOLEDB; User ID=sa;Password=SCL.2013.2013;data source=.\SQLEXPRESS;Initial Catalog=dbmas")
oConn.Open(MM_cnn_STRING)
Set DATOS = Server.CreateObject("ADODB.RecordSet")
DATOS.CursorType=3

	filtraRel=""
	if(vtipouser=1 and vrelator<>0)then
		 filtraRel=" and (bloque_programacion.id_relator='"&vrelator&"' or bloque_programacion.id_rel_seg='"&vrelator&"')"
	end if	
	
	if(vusr="144")then
		filtraRel=""
	end if

sql ="select distinct HISTORICO_CURSOS.RELATOR,"
sql = sql&"PROGRAMA.ID_PROGRAMA,CURRICULO.CODIGO,dbo.MayMinTexto(CURRICULO.NOMBRE_CURSO) as NOMBRE_CURSO,"
sql = sql&"CURRICULO.SENCE,CONVERT(VARCHAR(10),PROGRAMA.FECHA_INICIO_, 105) as FECHA_INICIO_,"
sql = sql&"(CASE WHEN bloque_programacion.id_rel_seg IS NOT NULL THEN dbo.MayMinTexto (INSTRUCTOR_RELATOR.NOMBRES+' '+INSTRUCTOR_RELATOR.A_PATERNO+' '+INSTRUCTOR_RELATOR.A_MATERNO) + ' / '+(select dbo.MayMinTexto (ri.NOMBRES+' '+ri.A_PATERNO+' '+ri.A_MATERNO) "
sql = sql&" from INSTRUCTOR_RELATOR ri where ri.ID_INSTRUCTOR=bloque_programacion.id_rel_seg) "
sql = sql&" else dbo.MayMinTexto (INSTRUCTOR_RELATOR.NOMBRES+' '+INSTRUCTOR_RELATOR.A_PATERNO+' '+INSTRUCTOR_RELATOR.A_MATERNO)  END) as instructor, "
sql = sql&" (case when CONVERT(date,PROGRAMA.FECHA_TERMINO, 105)<=CONVERT(date,GETDATE(),105) then "
sql = sql&" (case (select Datediff(""d"", Min(CONVERT(date,pg.FECHA_TERMINO, 105)), Max(CONVERT(date,GETDATE(),105))) " 
sql = sql&" from programa pg where pg.ID_PROGRAMA=PROGRAMA.ID_PROGRAMA) when 0 then '#66ff00' "
sql = sql&" when 1 then '#ffff00' else '#ff0000' end) else '#66ff00' end) as color, "
sql = sql&" CONVERT(date,PROGRAMA.FECHA_INICIO_, 105),PROGRAMA.DIR_EJEC,"
sql = sql&" (case when GETDATE()>=DATEADD(hour, 18, PROGRAMA.FECHA_TERMINO) then '1' else '0' end) as estadoEval, "
sql = sql&" HISTORICO_CURSOS.ID_BLOQUE, act_coordinador=isnull(estado_eva_cdn,0), "
sql = sql&" act_relator=isnull(estado_eva_rel,0),PROGRAMA.ID_MUTUAL,"
sql = sql&"cantPre=(SELECT COUNT(*) FROM HISTORICO_CURSOS HC WHERE HC.ID_BLOQUE=bloque_programacion.id_bloque AND hc.[PRE-CAL] is NOT null) from PROGRAMA "
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

fila=0
cieCursoCdn=0
WHILE NOT DATOS.EOF
	if(fila>=inicio AND fila<(limite*pagina))then
		cieCursoCdn=0
		Response.Write("<row id="""&fila&""">"&chr(13))
		Response.Write("<cell><![CDATA[<span style=""color:#000000; position:relative;"">"&DATOS("FECHA_INICIO_")&"<span style=""color:"&DATOS("color")&";top:-1px;left:-1px;position:absolute;"">"&DATOS("FECHA_INICIO_")&"</span></span>]]></cell>"&chr(13))					
		Response.Write("<cell><![CDATA["&DATOS("CODIGO")&"]]></cell>"&chr(13))
		Response.Write("<cell><![CDATA["&DATOS("NOMBRE_CURSO")&"]]></cell>"&chr(13))
		Response.Write("<cell><![CDATA["&DATOS("instructor")&"]]></cell>"&chr(13))
		Response.Write("<cell><![CDATA["&DATOS("DIR_EJEC")&"]]></cell>"&chr(13))		
		Response.Write("<cell><![CDATA[<span class=""ui-state-valid"" ><a href=""#"" title=""Libro de Clases"" class=""ui-icon ui-icon-folder-collapsed"" name=""aContrato"" onclick=""update("&DATOS("ID_PROGRAMA")&","&DATOS("RELATOR")&","&DATOS("ID_BLOQUE")&")""></a></span>]]></cell>"&chr(13))
		
		IF(DATOS("ID_MUTUAL")="73" or DATOS("ID_MUTUAL")="84" or DATOS("ID_MUTUAL")="85")THEN
			if(DATOS("cantPre")="0")then
				Response.Write("<cell><![CDATA[<span class=""ui-state-valid"" ><a href=""#"" title=""Pre-Evaluar Curso"" class=""ui-icon ui-icon-circle-check"" onclick=""evaCierrePre("&DATOS("ID_PROGRAMA")&","&DATOS("RELATOR")&","&vusr&","&DATOS("ID_BLOQUE")&",0,"&cieCursoCdn&")""></a></span>]]></cell>"&chr(13))		
			else
				Response.Write("<cell><![CDATA[<span class=""ui-state-valid"" ><a href=""#"" title=""Curso Pre-Evaluado"" class=""ui-icon ui-icon-circle-minus""></a></span>]]></cell>"&chr(13))	
			end if
		else
			Response.Write("<cell><![CDATA[<span class=""ui-state-valid"" ><a href=""#"" title=""Pre-Evaluaci&oacute;n no Disponible"" class=""ui-icon ui-icon-cancel"" name=""aContrato""></a></span>]]></cell>"&chr(13))	
		END IF
		
		
		if(DATOS("estadoEval")="1")then 
			cieCursoCdn=1
		end if
		
		       if(vtipouser=0 and DATOS("act_coordinador")="0")then
					'Response.Write("<cell><![CDATA[<span class=""ui-state-valid"" ><a href=""#"" title=""Evaluar Curso"" class=""ui-icon ui-icon-person"" onclick=""evaCierre2("&DATOS("ID_PROGRAMA")&","&DATOS("RELATOR")&","&vusr&","&DATOS("ID_BLOQUE")&",0,"&cieCursoCdn&")""></a></span>]]></cell>"&chr(13))
					Response.Write("<cell><![CDATA[<span class=""ui-state-valid"" ><a href=""#"" title=""Evaluar Curso"" class=""ui-icon ui-icon-person"" onclick=""evaCierre3("&DATOS("ID_PROGRAMA")&","&DATOS("RELATOR")&","&vusr&","&DATOS("ID_BLOQUE")&",0,"&cieCursoCdn&")""></a></span>]]></cell>"&chr(13))					
		       elseif(DATOS("act_coordinador")="1")then
					Response.Write("<cell><![CDATA[<span class=""ui-state-valid"" ><a href=""#"" title=""Curso Evaluado"" class=""ui-icon ui-icon-check""></a></span>]]></cell>"&chr(13))					
				else
					Response.Write("<cell><![CDATA[<span class=""ui-state-valid"" ><a href=""#"" title=""Disponible Solo Coordinador"" class=""ui-icon ui-icon-person""></a></span>]]></cell>"&chr(13))
				end if			
		
		if(DATOS("estadoEval")="1" or DATOS("ID_PROGRAMA")="12162" or DATOS("ID_PROGRAMA")="12167")then 
				if((vtipouser=1 and DATOS("act_relator")="0") or (vusr="143" and DATOS("act_relator")="0") or (vusr="8" and DATOS("act_relator")="0") or (vusr="36" and DATOS("act_relator")="0") or (vusr="108" and DATOS("act_relator")="0") or (vusr="72" and DATOS("act_relator")="0") or (vusr="172" and DATOS("act_relator")="0") or (vusr="32" and DATOS("act_relator")="0"))then
					Response.Write("<cell><![CDATA[<span class=""ui-state-valid"" ><a href=""#"" title=""Evaluar Curso"" class=""ui-icon ui-icon-person"" onclick=""evaCierre("&DATOS("ID_PROGRAMA")&","&DATOS("RELATOR")&","&vusr&","&DATOS("ID_BLOQUE")&",1)""></a></span>]]></cell>"&chr(13))
				elseif(DATOS("act_relator")="1")then
					Response.Write("<cell><![CDATA[<span class=""ui-state-valid"" ><a href=""#"" title=""Curso Evaluado"" class=""ui-icon ui-icon-check""></a></span>]]></cell>"&chr(13))					
				else
					Response.Write("<cell><![CDATA[<span class=""ui-state-valid"" ><a href=""#"" title=""Disponible Solo Relator"" class=""ui-icon ui-icon-person""></a></span>]]></cell>"&chr(13))
				end if				
		else
				'Response.Write("<cell><![CDATA[<span class=""ui-state-valid"" ><a href=""#"" title=""Curso En Proceso"" class=""ui-icon ui-icon-cancel"" name=""aContrato""></a></span>]]></cell>"&chr(13))	
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
