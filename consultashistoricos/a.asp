<!--#include file="../conexion.asp"-->
<%
Response.CodePage = 65001
Response.CharSet = "utf-8"
Response.ContentType="text/xml"
Response.Write("<?xml version='1.0' encoding='UTF-8'?>")
'Response.Write("<modificar>") 

'on error resume next

'query = "update facturas set factura=6312,FECHA_EMISION=convert(date, '26-05-2015') where id_factura=39923 and factura is null;"
'query = "update facturas set factura=6313,FECHA_EMISION=convert(date, '26-05-2015') where id_factura=39922 and factura is null;"
'query = "update facturas set factura=6314,FECHA_EMISION=convert(date, '26-05-2015') where id_factura=37715 and factura is null;"
'query = "update facturas set factura=6315,FECHA_EMISION=convert(date, '26-05-2015') where id_factura=38088 and factura is null;"
'query = "update facturas set factura=6316,FECHA_EMISION=convert(date, '26-05-2015') where id_factura=38102 and factura is null;"

'response.Write(query)
'response.End()

'conn.execute (query)
'Response.Write("<sql>"&query&"</sql>")
'if err.number <> 0 then
'	Response.Write("<commit>false</commit>")
'else
'	Response.Write("<commit>true</commit>")
'end if
'Response.Write("</modificar>")

qsHc="select f.factura,A.ORDEN_COMPRA,a.id_autorizacion from facturas f inner join AUTORIZACION a on a.id_autorizacion=f.id_autorizacion where A.ORDEN_COMPRA in('Convenio') and f.id_factura=40335"

set rsHc =  conn.execute(qsHc) 
fila=1
Response.Write("<rows>"&chr(13)) 

while not rsHc.eof
		Response.Write("<row id="""&fila&""">"&chr(13))
		Response.Write("<cell><![CDATA["&rsHc("ORDEN_COMPRA")&"]]></cell>"&chr(13))
		Response.Write("<cell><![CDATA["&rsHc("factura")&"]]></cell>"&chr(13))
		Response.Write("<cell><![CDATA["&rsHc("id_autorizacion")&"]]></cell>"&chr(13))
		Response.Write("</row>"&chr(13))
  		rsHc.MoveNext
   wend
Response.Write("</rows>") 
%>