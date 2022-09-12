<%Response.CodePage = 65001
Response.CharSet = "utf-8"
Response.ContentType="text/xml"
Response.Write("<?xml version='1.0' encoding='UTF-8'?>")
Response.Write("<insertar>") %>
<!--#include file="../conexion.asp"-->
<%
on error resume next
vmovbdg = Request("MovBdg")
vmovfecha = Request("MovFecha")
vmovart = Request("MovArt")
vmovtipo = Request("MovTipo")
vmovcant = Request("MovCantidad")

vmovtipodoc = Request("MovTipoDoc")
vmovfechadoc = Request("MovFechaDoc")

'vmovprecio = replace(Request("MovPrecio"),".",",")
vmovprecio = Request("MovPrecio")

dim query
query = "insert into MOVIMIENTOS (ID_ARTICULO,TIPO_MOVIMIENTO,CANTIDAD,ESTADO,FECHA,ID_BODEGA,MODULO,ID_USUARIO,TIPO_DOCUMENTO,"
query = query&"FECHA_DOC) values('"&vmovart&"','"&vmovtipo&"','"&vmovcant&"',1,CONVERT(datetime,'"&vmovfecha&"',105),"
query = query&"'"&vmovbdg&"',1,'"&Session("usuarioMutual")&"','"&vmovtipodoc&"',CONVERT(datetime,'"&vmovfechadoc&"',105)) "

conn.execute (query)

dim upArt
upArt="update ARTICULO_BODEGA set FECHA_ULT_ENTRADA=GETDATE(),ULT_VALOR_UNITARIO=CONVERT(FLOAT,'"&vmovprecio&"')"
upArt = upArt&",STOCK_ACTUAL=(select STOCK_ACTUAL from ARTICULO_BODEGA where ESTADO_ARTICULO_BODEGA=1 "
upArt = upArt&" AND ID_ARTICULO IS NOT NULL AND ID_ARTICULO='"&vmovart&"' and ID_BODEGA='"&vmovbdg&"')+"&vmovcant
upArt = upArt&" WHERE (ESTADO_ARTICULO_BODEGA = 1) AND (ID_ARTICULO IS NOT NULL) "
upArt = upArt&" and ID_ARTICULO='"&vmovart&"' and ID_BODEGA='"&vmovbdg&"'"

conn.execute (upArt)

if err.number <> 0 then
	Response.Write("<commit>false</commit>")
else
	Response.Write("<commit>true</commit>")
end if
Response.Write("</insertar>") 
%>