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
columna = "INSTRUCTOR_RELATOR.RUT"
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

sql = "select INSTRUCTOR_RELATOR.RUT, "
sql = sql&"dbo.MayMinTexto (INSTRUCTOR_RELATOR.NOMBRES+' '+INSTRUCTOR_RELATOR.A_PATERNO+' '+INSTRUCTOR_RELATOR.A_MATERNO) "
sql = sql&"as instructor,(CASE WHEN bloqProg.id_sede =  27 THEN bloqProg.nom_sede " 
sql = sql&"WHEN bloqProg.id_sede <>  27 THEN SEDES.NOMBRE+', '+SEDES.DIRECCION+', '+SEDES.CIUDAD END) as sede, "
sql = sql&"(CASE WHEN bloqProg.id_rel_seg IS NOT NULL THEN ' / '+(select dbo.MayMinTexto (ri.NOMBRES+' '"
sql = sql&"+ri.A_PATERNO+' '+ri.A_MATERNO) from INSTRUCTOR_RELATOR ri where ri.ID_INSTRUCTOR=bloqProg.id_rel_seg) " 
sql = sql&"WHEN bloqProg.id_rel_seg IS NULL THEN '' END) as rel_seg, "
sql = sql&"bloqProg.id_bloque,bloqProg.cupos,bloqProg.estado,"
sql = sql&"(select COUNT(*) from HISTORICO_CURSOS where HISTORICO_CURSOS.ID_BLOQUE=bloqProg.id_bloque) as incritos "
sql = sql&" from bloque_programacion bloqProg"
sql = sql&" inner join SEDES on SEDES.ID_SEDE=bloqProg.id_sede "
sql = sql&" inner join INSTRUCTOR_RELATOR on INSTRUCTOR_RELATOR.ID_INSTRUCTOR=bloqProg.id_relator "
sql = sql&" where bloqProg.id_programa='"&request("id_prog")&"'"
'sql = sql&" ORDER BY "&columna&" "&orden

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

'response.Write(sql)
'response.end()

fila=0
WHILE NOT DATOS.EOF
	if(fila>=inicio AND fila<(limite*pagina))then
		Response.Write("<row id="""&fila&""">"&chr(13))
		'Response.Write("<cell><![CDATA["&replace(FormatNumber(mid(DATOS("RUT"), 1,len(DATOS("RUT"))-2),0)&mid(DATOS("RUT"), len(DATOS("RUT"))-1,len(DATOS("RUT"))),",",".")&"]]></cell>"&chr(13))
		Response.Write("<cell><![CDATA["&DATOS("instructor")&DATOS("rel_seg")&"]]></cell>"&chr(13))
		Response.Write("<cell><![CDATA["&DATOS("sede")&"]]></cell>"&chr(13))
		Response.Write("<cell><![CDATA["&DATOS("cupos")&"]]></cell>"&chr(13))
		Response.Write("<cell><![CDATA["&DATOS("incritos")&"]]></cell>"&chr(13))
		Response.Write("<cell><![CDATA[<span class=""ui-state-valid"" ><a href=""#"" title=""Modificar Registro"" class=""ui-icon ui-icon-pencil"" name=""aContrato"" onclick=""update_tab("&request("id_prog")&","&DATOS("id_bloque")&")""></a></span>]]></cell>"&chr(13))

		if(DATOS("estado")="0")then
		Response.Write("<cell><![CDATA[<span class=""ui-state-valid"" ><a href=""#"" title=""Abrir Bloque"" class=""ui-icon ui-icon-circle-close"" name=""aContrato"" onclick=""cerrar_bloque("&request("id_prog")&","&DATOS("id_bloque")&","&DATOS("estado")&")""></a></span>]]></cell>"&chr(13))
		else
		Response.Write("<cell><![CDATA[<span class=""ui-state-valid"" ><a href=""#"" title=""Cerrar Bloque"" class=""ui-icon ui-icon-circle-check"" name=""aContrato"" onclick=""cerrar_bloque("&request("id_prog")&","&DATOS("id_bloque")&","&DATOS("estado")&")""></a></span>]]></cell>"&chr(13))
		end if
		
		if(cdbl(DATOS("incritos"))>0)then
		Response.Write("<cell><![CDATA[<span class=""ui-state-valid"" ><a href=""#"" title=""No se puede Eliminar el Bloque"" class=""ui-icon ui-icon-minusthick"" name=""aContrato"" ></a></span>]]></cell>"&chr(13))
		else
		Response.Write("<cell><![CDATA[<span class=""ui-state-valid"" ><a href=""#"" title=""Eliminar Bloque"" class=""ui-icon ui-icon-trash"" name=""aContrato"" onclick=""eliminarBloque("&request("id_prog")&","&DATOS("id_bloque")&")""></a></span>]]></cell>"&chr(13))
		end if
		
		Response.Write("<cell><![CDATA[<span class=""ui-state-valid"" ><a href=""#"" title=""Ingreso Costos"" class=""ui-icon ui-icon-calculator"" name=""aContrato"" onclick=""ing_gastos("&DATOS("id_bloque")&")""></a></span>]]></cell>"&chr(13))

		Response.Write("</row>"&chr(13))
	end if
	fila=fila+1
	DATOS.MoveNext
WEND
Response.Write("</rows>") 
%>
