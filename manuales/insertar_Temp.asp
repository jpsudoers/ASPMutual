<%Response.CodePage = 65001
Response.CharSet = "utf-8"
Response.ContentType="text/xml"
Response.Write("<?xml version='1.0' encoding='UTF-8'?>")
Response.Write("<insertar>") %>
<!--#include file="../conexion.asp"-->
<%
on error resume next
vbdgart = Request("BdgArt")
vidprov = Request("IdProvAsig")
vcantmin = Request("CantMinArt")
vstockcrit = Request("StockCritArt")
vcantrep = Request("CantRepArt")

dim query
if(Request("TipoInsert")="0")then
query = "insert into ARTICULO_BODEGA (ID_BODEGA,ID_ARTICULO_PROV,ESTADO_ARTICULO_BODEGA,ART_MINIMOS,STOCK_CRITICO,REP_ARTICULO,"
query = query&"ESTADO_REGISTRO,STOCK_ACTUAL) values('"&vbdgart&"','"&vidprov&"',1,'"&vcantmin&"','"&vstockcrit&"',"
query = query&"'"&vcantrep&"',1,0)"
else
query = "insert into ARTICULO_BODEGA (ID_BODEGA,ID_ARTICULO,ESTADO_ARTICULO_BODEGA,ART_MINIMOS,STOCK_CRITICO,REP_ARTICULO,"
query = query&"ESTADO_REGISTRO,STOCK_ACTUAL) values('"&vbdgart&"','"&vidprov&"',1,'"&vcantmin&"',"
query = query&"'"&vstockcrit&"','"&vcantrep&"',1,0)"
end if

conn.execute (query)

if err.number <> 0 then
	Response.Write("<commit>false</commit>")
else
	Response.Write("<commit>true</commit>")
end if
Response.Write("</insertar>") 
%>