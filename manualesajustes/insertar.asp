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
vmovtipoajuste = Request("MovTipoAjusteSel")
vmovexpli = Request("MovExpli")

dim signo
if(vmovtipoajuste="1")then
	signo="+"
else
	signo="-"
end if

dim query
query = "insert into MOVIMIENTOS (ID_ARTICULO,TIPO_MOVIMIENTO,CANTIDAD,ESTADO,FECHA,ID_BODEGA,MODULO,TIPO_AJUSTE,EXPLICACION"
query = query&",ID_USUARIO) values('"&vmovart&"','"&vmovtipo&"','"&signo&vmovcant&"',1,CONVERT(datetime,'"&vmovfecha&"',105),"
query = query&"'"&vmovbdg&"',3,'"&vmovtipoajuste&"','"&vmovexpli&"','"&Session("usuarioMutual")&"') "

conn.execute (query)

dim upArt
upArt = "update ARTICULO_BODEGA set STOCK_ACTUAL="
upArt = upArt&"(select STOCK_ACTUAL from ARTICULO_BODEGA where ESTADO_ARTICULO_BODEGA=1 "
upArt = upArt&" AND ID_ARTICULO IS NOT NULL AND ID_ARTICULO='"&vmovart&"' and ID_BODEGA='"&vmovbdg&"')"&signo&vmovcant
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