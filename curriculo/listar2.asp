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
columna = "descripcion"
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

  sql="select cb.ID_MUTUAL,cb.NOMBRE_BLOQUE,cb.horas,cb.DIA,ca.TEMA,ca.ACTIVIDAD,ca.ID_ACTIVIDAD_CURSO,ca.ID_BLOQUE_CURSO"&_
	  " from CURRICULO_BLOQUE cb "&_
	  " inner join CURRICULO_ACTIVIDADES ca on ca.ID_BLOQUE_CURSO=cb.ID_BLOQUE_CURSO "&_
	  " where cb.ID_MUTUAL="&Request("id")&" and ca.ESTADO=1 and cb.ESTADO=1 "&_
	  " order by cb.DIA,cb.NOMBRE_BLOQUE "
	  '&columna&" "&orden

'response.Write(sql)
'response.end()

DATOS.Open sql,oConn
total_pages = 1
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
WHILE NOT DATOS.EOF
	if(fila>=inicio AND fila<(limite*pagina))then
		Response.Write("<row id="""&fila&""">"&chr(13))
		Response.Write("<cell><![CDATA["&DATOS("DIA")&"]]></cell>"&chr(13))
		Response.Write("<cell><![CDATA["&DATOS("NOMBRE_BLOQUE")&"]]></cell>"&chr(13))
		Response.Write("<cell><![CDATA["&DATOS("TEMA")&"]]></cell>"&chr(13))
		Response.Write("<cell><![CDATA["&DATOS("ACTIVIDAD")&"]]></cell>"&chr(13))		
		Response.Write("<cell><![CDATA["&DATOS("horas")&"]]></cell>"&chr(13))		
		Response.Write("<cell><![CDATA[<span class=""ui-state-valid""><a href=""#"" title=""Modificar Actividad"" class=""ui-icon ui-icon-pencil"" onclick=""updateAct("&DATOS("ID_ACTIVIDAD_CURSO")&","&DATOS("ID_MUTUAL")&")""></a></span>]]></cell>"&chr(13))
		Response.Write("<cell><![CDATA[<span class=""ui-state-valid"" ><a href=""#"" title=""Eliminar Actividad"" class=""ui-icon ui-icon-trash"" onclick=""eliminarAct("&DATOS("ID_ACTIVIDAD_CURSO")&","&DATOS("ID_ACTIVIDAD_CURSO")&")""></a></span>]]></cell>"&chr(13))
		Response.Write("</row>"&chr(13))
	end if
	fila=fila+1
	DATOS.MoveNext
WEND
Response.Write("</rows>") 
%>
