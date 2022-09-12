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
columna = "B.UBICACION"
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

sql="SELECT AB.ID_ARTICULO_BODEGA,AB.ID_ARTICULO_PROV,AB.ART_MINIMOS,AB.STOCK_CRITICO,AB.STOCK_ACTUAL,AB.REP_ARTICULO,B.UBICACION,"
sql = sql&" (B.UBICACION+', '+B.DIRECCION) AS NOM_BODEGA FROM ARTICULO_BODEGA AB "
sql = sql&" inner join BODEGAS B on B.ID_BODEGA=AB.ID_BODEGA "

if(request("tipoBusqueda")="0")then
sql = sql&" where AB.ESTADO_ARTICULO_BODEGA=1 and AB.ID_ARTICULO_PROV='"&request("IdArt")&"' and ESTADO_REGISTRO<>'"&Request("est")&"'"
else
sql = sql&" where AB.ESTADO_ARTICULO_BODEGA=1 and AB.ID_ARTICULO='"&request("IdArt")&"' and ESTADO_REGISTRO<>'"&Request("est")&"'"
end if

sql = sql&" ORDER BY "&columna&" "&orden

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
		Response.Write("<cell><![CDATA["&DATOS("NOM_BODEGA")&"]]></cell>"&chr(13))
		
		if(cdbl(DATOS("STOCK_ACTUAL"))<=cdbl(DATOS("ART_MINIMOS")))then
		     if(cdbl(DATOS("STOCK_ACTUAL"))<=cdbl(DATOS("STOCK_CRITICO")))then
			   Response.Write("<cell><![CDATA[<font color=""red""><b>"&DATOS("STOCK_ACTUAL")&"</b></font>]]></cell>"&chr(13))
			 else
			   Response.Write("<cell><![CDATA[<font color=""#3333cc""><b>"&DATOS("STOCK_ACTUAL")&"</b></font>]]></cell>"&chr(13))
			 end if
		else
			Response.Write("<cell><![CDATA[<b>"&DATOS("STOCK_ACTUAL")&"</b>]]></cell>"&chr(13))
		end if
				
		Response.Write("<cell><![CDATA["&DATOS("ART_MINIMOS")&"]]></cell>"&chr(13))
		Response.Write("<cell><![CDATA["&DATOS("STOCK_CRITICO")&"]]></cell>"&chr(13))
		Response.Write("<cell><![CDATA["&DATOS("REP_ARTICULO")&"]]></cell>"&chr(13))
		Response.Write("<cell><![CDATA[<span class=""ui-state-valid"" ><a href=""#"" title=""Modificar Registro"" class=""ui-icon ui-icon-pencil"" name=""aContrato"" onclick=""updateBdg("&DATOS("ID_ARTICULO_BODEGA")&")""></a></span>]]></cell>"&chr(13))
		Response.Write("<cell><![CDATA[<span class=""ui-state-valid"" ><a href=""#"" title=""Eliminar Registro"" class=""ui-icon ui-icon-trash"" name=""aContrato"" onclick=""eliminarBdg("&DATOS("ID_ARTICULO_BODEGA")&")""></a></span>]]></cell>"&chr(13))	
		Response.Write("</row>"&chr(13))
	end if
	fila=fila+1
	DATOS.MoveNext
WEND
Response.Write("</rows>") 
%>
