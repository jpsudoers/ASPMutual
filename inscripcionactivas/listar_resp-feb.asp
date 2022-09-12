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
columna = "PROGRAMA.FECHA_INICIO_"
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

prog=""
if(Request("programa")<>"0")then
prog=" and AUTORIZACION.ID_PROGRAMA in (select p.ID_PROGRAMA from PROGRAMA p "
prog=prog&" where CONVERT(date,p.FECHA_INICIO_,105)=CONVERT(date,'"&Request("programa")&"',105))"
end if

emp=""
if(Request("empresa")<>"")then
emp=" and AUTORIZACION.ID_EMPRESA='"&Request("empresa")&"'"
end if

TipoDoc=""
if(Request("doc")<>"T")then
	if(Request("doc")="SC")then
		TipoDoc=" and AUTORIZACION.TIPO_DOC<>4"
	else
		TipoDoc=" and AUTORIZACION.TIPO_DOC=4"
	end if
end if

sql = "select EMPRESAS.RUT as rut,dbo.MayMinTexto (EMPRESAS.R_SOCIAL) as empresa, "
sql = sql&"CONVERT(VARCHAR(10),PROGRAMA.FECHA_INICIO_, 105) as FECHA_INICIO, "
sql = sql&" dbo.MayMinTexto (CURRICULO.NOMBRE_CURSO) as nombre, "
sql = sql&" AUTORIZACION.ID_AUTORIZACION,AUTORIZACION.ID_PROGRAMA as ID_PROGRAMA,AUTORIZACION.DOCUMENTO_COMPROMISO as doc,"
sql = sql&" dbo.MayMinTexto (INSTRUCTOR_RELATOR.NOMBRES+' '+INSTRUCTOR_RELATOR.A_PATERNO) as instructor,"
sql = sql&"(CASE WHEN AUTORIZACION.CON_OTIC=1 AND AUTORIZACION.CON_FRANQUICIA=1 AND AUTORIZACION.TIPO_DOC=4 "
sql = sql&" AND (AUTORIZACION.N_REG_SENCE IS NULL OR AUTORIZACION.N_REG_SENCE IS NOT NULL) then '0' " 
sql = sql&"  WHEN AUTORIZACION.CON_OTIC=0 AND AUTORIZACION.CON_FRANQUICIA=1 AND AUTORIZACION.TIPO_DOC=4 AND AUTORIZACION.N_REG_SENCE IS NOT NULL then '0' " 
sql = sql&"  WHEN AUTORIZACION.CON_OTIC=0 AND AUTORIZACION.CON_FRANQUICIA=0 AND AUTORIZACION.TIPO_DOC=4 then '0' "
'sql = sql&"  WHEN AUTORIZACION.CON_OTIC=1 AND AUTORIZACION.CON_FRANQUICIA=1 AND AUTORIZACION.TIPO_DOC<>4 AND AUTORIZACION.N_REG_SENCE IS NULL then '0' " 
sql = sql&"  WHEN AUTORIZACION.CON_OTIC=0 AND AUTORIZACION.CON_FRANQUICIA=1 AND AUTORIZACION.TIPO_DOC<>4 then '1' " 
sql = sql&"  WHEN AUTORIZACION.CON_OTIC=0 AND AUTORIZACION.CON_FRANQUICIA=0 AND AUTORIZACION.TIPO_DOC<>4 then '1' "  
sql = sql&"  WHEN AUTORIZACION.CON_OTIC=1 AND AUTORIZACION.CON_FRANQUICIA=1 AND AUTORIZACION.TIPO_DOC<>4 AND AUTORIZACION.N_REG_SENCE IS NOT NULL then '1' END) as 'ConOticyCarta', "
sql = sql&" case when (select Datediff(""d"", Min(CONVERT(date,GETDATE(),105)), Max(CONVERT(date,pg.FECHA_TERMINO, 105))) "  
sql = sql&" from programa pg where pg.ID_PROGRAMA=PROGRAMA.ID_PROGRAMA)<=(-2) then '#ff0000' "  
sql = sql&" when (select Datediff(""d"", Min(CONVERT(date,GETDATE(),105)), Max(CONVERT(date,pg.FECHA_TERMINO, 105))) "  
sql = sql&" from programa pg where pg.ID_PROGRAMA=PROGRAMA.ID_PROGRAMA)=(-1) then '#ffff00' " 
sql = sql&" when (select Datediff(""d"", Min(CONVERT(date,GETDATE(),105)), Max(CONVERT(date,pg.FECHA_TERMINO, 105))) "  
sql = sql&" from programa pg where pg.ID_PROGRAMA=AUTORIZACION.ID_PROGRAMA)>=0 then '#66ff00' end as color  "
sql = sql&" from AUTORIZACION "
sql = sql&" inner join PROGRAMA on PROGRAMA.ID_PROGRAMA=AUTORIZACION.ID_PROGRAMA " 
sql = sql&" inner join EMPRESAS on EMPRESAS.ID_EMPRESA=AUTORIZACION.ID_EMPRESA " 
sql = sql&" inner join CURRICULO on CURRICULO.ID_MUTUAL=PROGRAMA.ID_MUTUAL "
sql = sql&" inner join bloque_programacion on bloque_programacion.id_bloque=AUTORIZACION.ID_BLOQUE "
sql = sql&" inner join INSTRUCTOR_RELATOR on INSTRUCTOR_RELATOR.ID_INSTRUCTOR=bloque_programacion.id_relator "
sql = sql&" WHERE AUTORIZACION.ESTADO=1 "&prog&emp&TipoDoc&" ORDER BY "&columna&" "&orden

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
		Response.Write("<cell><![CDATA[<span style=""color:#000000; position:relative;"">"&DATOS("FECHA_INICIO")&"<span style=""color:"&DATOS("color")&";top:-1px;left:-1px;position:absolute;"">"&DATOS("FECHA_INICIO")&"</span></span>]]></cell>"&chr(13))		
		Response.Write("<cell><![CDATA["&DATOS("rut")&"]]></cell>"&chr(13))
		Response.Write("<cell><![CDATA["&DATOS("empresa")&"]]></cell>"&chr(13))
		Response.Write("<cell><![CDATA["&DATOS("nombre")&"]]></cell>"&chr(13))
		Response.Write("<cell><![CDATA["&DATOS("instructor")&"]]></cell>"&chr(13))
		

		if(DATOS("ConOticyCarta")="0")then
		Response.Write("<cell><![CDATA[<span><center><img src=""images/rojo.png"" width=""13"" height=""13""/></center></span>]]></cell>"&chr(13))
		else
		Response.Write("<cell><![CDATA[<span><center><img src=""images/revisado.png"" width=""13"" height=""13""/></center></span>]]></cell>"&chr(13))
		end if
		
		Response.Write("<cell><![CDATA[<span class=""ui-state-valid"" ><a href=""#"" title=""Ver Documento"" class=""ui-icon ui-icon-document"" name=""aContrato"" onclick=""documento('"&DATOS("doc")&"');""></a></span>]]></cell>"&chr(13))
		Response.Write("<cell><![CDATA[<span class=""ui-state-valid"" ><a href=""#"" title=""Modificar Registro"" class=""ui-icon ui-icon-pencil"" name=""aContrato"" onclick=""update("&DATOS("ID_AUTORIZACION")&")""></a></span>]]></cell>"&chr(13))
		Response.Write("<cell><![CDATA[<span class=""ui-state-valid"" ><a href=""#"" title=""Eliminar Registro"" class=""ui-icon ui-icon-trash"" name=""aContrato"" onclick=""delInscripcion('"&DATOS("ID_AUTORIZACION")&"');""></a></span>]]></cell>"&chr(13))
		Response.Write("</row>"&chr(13))
	end if
	fila=fila+1
	DATOS.MoveNext
WEND
Response.Write("</rows>") 
%>
