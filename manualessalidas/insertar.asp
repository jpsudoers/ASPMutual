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
vmovcodcurso = Request("MovCodCurso")

dim query
query = "insert into MOVIMIENTOS (ID_ARTICULO,TIPO_MOVIMIENTO,CANTIDAD,ESTADO,FECHA,ID_BODEGA,MODULO,ID_PROG_BLOQUE,ID_USUARIO)"
query = query&" values('"&vmovart&"',3,'"&"-"&vmovcant&"',1,CONVERT(datetime,'"&vmovfecha&"',105),"
query = query&"'"&vmovbdg&"',2,'"&vmovcodcurso&"','"&Session("usuarioMutual")&"') "

conn.execute (query)

dim upArt
upArt = "update ARTICULO_BODEGA set FECHA_ULT_SALIDA=GETDATE(),STOCK_ACTUAL="
upArt = upArt&"(select STOCK_ACTUAL from ARTICULO_BODEGA where ESTADO_ARTICULO_BODEGA=1 "
upArt = upArt&" AND ID_ARTICULO IS NOT NULL AND ID_ARTICULO='"&vmovart&"' and ID_BODEGA='"&vmovbdg&"')-"&vmovcant
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