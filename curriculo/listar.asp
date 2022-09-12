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

sql = "SELECT ID_MUTUAL,CODIGO,(CASE SENCE WHEN 0 THEN 'Si' else 'No' END) as SENCE,dbo.MayMinTexto(NOMBRE_CURSO) as NOMBRE_CURSO from CURRICULO WHERE ESTADO=1 ORDER BY "&columna&" "&orden

'response.Write(sql)
'response.end()

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
WHILE NOT DATOS.EOF
	if(fila>=inicio AND fila<(limite*pagina))then
		Response.Write("<row id="""&fila&""">"&chr(13))
		Response.Write("<cell><![CDATA["&DATOS("CODIGO")&"]]></cell>"&chr(13))
		Response.Write("<cell><![CDATA["&DATOS("SENCE")&"]]></cell>"&chr(13))
		Response.Write("<cell><![CDATA["&DATOS("NOMBRE_CURSO")&"]]></cell>"&chr(13))
		Response.Write("<cell><![CDATA[<span class=""ui-state-valid"" ><a href=""#"" title=""Ingreso Actividades"" class=""ui-icon ui-icon-comment"" onclick=""detActividades("&DATOS("ID_MUTUAL")&")""></a></span>]]></cell>"&chr(13))		
		Response.Write("<cell><![CDATA[<span class=""ui-state-valid""><a href=""#"" title=""Modificar Registro"" class=""ui-icon ui-icon-pencil"" name=""aContrato"" onclick=""update("&DATOS("ID_MUTUAL")&")""></a></span>]]></cell>"&chr(13))
		Response.Write("<cell><![CDATA[<span class=""ui-state-valid"" ><a href=""#"" title=""Eliminar Registro"" class=""ui-icon ui-icon-trash"" name=""aContrato"" onclick=""eliminar("&DATOS("ID_MUTUAL")&")""></a></span>]]></cell>"&chr(13))
		Response.Write("</row>"&chr(13))
	end if
	fila=fila+1
	DATOS.MoveNext
WEND
Response.Write("</rows>") 
%>
