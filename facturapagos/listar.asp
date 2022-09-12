<!--#include file="../cnn_string.asp"-->
<!--#include file="../conexion.asp"-->
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
columna = "FC.ID_FACTURA"
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

sql = "select E.RUT as rut,dbo.MayMinTexto(E.R_SOCIAL) as empresa,FC.MONTO,FC.ID_FACTURA,FC.FACTURA,"
sql = sql&" (Select Datediff(""d"", Min(FC2.FECHA_VENCIMIENTO), Max(CONVERT(date,GETDATE(), 105))) " 
sql = sql&" from FACTURAS FC2 where FC2.ID_FACTURA=FC.ID_FACTURA) as dias, "
sql = sql&"(CASE WHEN FC.TIPO_DOC='0' then 'Orden de Compra N° '+FC.N_DOCUMENTO " 
sql = sql&" WHEN FC.TIPO_DOC='1' then 'Vale Vista N° '+FC.N_DOCUMENTO "  
sql = sql&" WHEN FC.TIPO_DOC='2' then 'Depósito Cheque N° '+FC.N_DOCUMENTO " 
sql = sql&" WHEN FC.TIPO_DOC='3' then 'Transferencia N° ' +FC.N_DOCUMENTO " 
sql = sql&" WHEN FC.TIPO_DOC='4' then 'Carta Compromiso N° ' +FC.N_DOCUMENTO " 
sql = sql&" END) as 'Documento' from FACTURAS FC " 
sql = sql&" inner join EMPRESAS E on E.ID_EMPRESA=FC.ID_EMPRESA " 
sql = sql&" where FC.ESTADO_CANCELACION=0 " 
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
imagen=""
WHILE NOT DATOS.EOF
	imagen="<cell><![CDATA[<span></span>]]></cell>"
	if(fila>=inicio AND fila<(limite*pagina))then
		Response.Write("<row id="""&fila&""">"&chr(13))
		Response.Write("<cell><![CDATA["&DATOS("rut")&"]]></cell>"&chr(13))
		Response.Write("<cell><![CDATA["&DATOS("empresa")&"]]></cell>"&chr(13))  
		Response.Write("<cell><![CDATA["&DATOS("Documento")&"]]></cell>"&chr(13))
		Response.Write("<cell><![CDATA["&DATOS("FACTURA")&"]]></cell>"&chr(13))
		Response.Write("<cell><![CDATA["&"$ "&replace(FormatNumber(DATOS("MONTO"),0),",",".")&"]]></cell>"&chr(13))
		Response.Write("<cell><![CDATA[<span class=""ui-state-valid""><a href=""#"" title=""Modificar Registro"" class=""ui-icon ui-icon-pencil"" name=""aContrato"" onclick=""update("&DATOS("ID_FACTURA")&")""></a></span>]]></cell>"&chr(13))
		
		if(DATOS("dias")>0 and DATOS("dias")<=15)then
			imagen="<cell><![CDATA[<span><center><img src=""images/amarillo.png"" width=""16"" height=""16""/></center></span>]]></cell>"
		end if
		
		if(DATOS("dias")>15 and DATOS("dias")<=30)then
			imagen="<cell><![CDATA[<span><center><img src=""images/naranja.png"" width=""16"" height=""16""/></center></span>]]></cell>"
		end if
		
		if(DATOS("dias")>30)then
			imagen="<cell><![CDATA[<span><center><img src=""images/rojo.png"" width=""16"" height=""16""/></center></span>]]></cell>"
		end if
		
		Response.Write(imagen&chr(13))
		Response.Write("</row>"&chr(13))
	end if
	fila=fila+1
	imagen=""
	DATOS.MoveNext
WEND
Response.Write("</rows>") 
%>

