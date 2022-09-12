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
columna = "id_ingresos"
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

sql = "select monto,CONVERT(VARCHAR(10),fecha_pago, 105) as fecha,comprobante,(CASE WHEN forma_pago ='1' then 'Vale Vista' "
sql = sql&" WHEN forma_pago ='2' then 'Depósito Cheques' WHEN forma_pago ='3' then 'Transferencia Electrónica' " 
sql = sql&" WHEN forma_pago ='5' then 'Cheque a Fecha' END) as 'formapago' from INGRESOS " 
sql = sql&" where id_factura='"&Request("id_factura")&"' order by id_ingresos asc " 

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
		Response.Write("<cell><![CDATA["&DATOS("fecha")&"]]></cell>"&chr(13))
		Response.Write("<cell><![CDATA["&DATOS("formapago")&"]]></cell>"&chr(13))
		Response.Write("<cell><![CDATA["&DATOS("comprobante")&"]]></cell>"&chr(13))
		Response.Write("<cell><![CDATA["&"$ "&replace(FormatNumber(DATOS("monto"),0),",",".")&"]]></cell>"&chr(13))		
		Response.Write("</row>"&chr(13))
	end if
	fila=fila+1
	DATOS.MoveNext
WEND
Response.Write("</rows>") 
%>
