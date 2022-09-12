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
orden = "DESC"
if(Request("sord") <> "")then orden = Request("sord")
columna = "M.FECHA"
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

sql = "select CONVERT(VARCHAR(10),M.FECHA, 105) as FECHA,TM.NOMBRE_MOV as TIPO_MOVIMIENTO,"
sql = sql&" M.CANTIDAD,B.UBICACION+', '+B.DIRECCION AS NOM_BODEGA,B.UBICACION,A.DESC_ARTICULO "
sql = sql&" FROM MOVIMIENTOS M "
sql = sql&" inner join articulos A on A.id_articulo=M.id_articulo "
sql = sql&" inner join BODEGAS B on B.ID_BODEGA=M.ID_BODEGA "
sql = sql&" inner join TIPO_MOVIMIENTO TM on TM.ID_TIPO_MOV=M.TIPO_MOVIMIENTO "
sql = sql&" where M.MODULO=1 AND M.ESTADO=1 "
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
		Response.Write("<cell><![CDATA["&DATOS("FECHA")&"]]></cell>"&chr(13))
		Response.Write("<cell><![CDATA["&DATOS("TIPO_MOVIMIENTO")&"]]></cell>"&chr(13))		
		Response.Write("<cell><![CDATA["&DATOS("DESC_ARTICULO")&"]]></cell>"&chr(13))
		Response.Write("<cell><![CDATA["&DATOS("NOM_BODEGA")&"]]></cell>"&chr(13))		
		Response.Write("<cell><![CDATA["&DATOS("CANTIDAD")&"]]></cell>"&chr(13))
		Response.Write("</row>"&chr(13))
	end if
	fila=fila+1
	DATOS.MoveNext
WEND
Response.Write("</rows>") 
%>
