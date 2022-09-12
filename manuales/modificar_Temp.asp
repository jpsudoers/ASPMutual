<!--#include file="../conexion.asp"-->
<%
Response.CodePage = 65001
Response.CharSet = "utf-8"
Response.ContentType="text/xml"
Response.Write("<?xml version='1.0' encoding='UTF-8'?>")
Response.Write("<modificar>") 

on error resume next
vId = Request("ID_Art_Bdg")
vbdgart = Request("BdgArt")
vcantmin = Request("CantMinArt")
vstockcrit = Request("StockCritArt")
vcantrep = Request("CantRepArt")

query = "UPDATE ARTICULO_BODEGA SET ID_BODEGA='"&vbdgart&"', ART_MINIMOS='"&vcantmin&"',"
query = query&" STOCK_CRITICO='"&vstockcrit&"', REP_ARTICULO='"&vcantrep&"' "
query = query&" WHERE ID_ARTICULO_BODEGA='"&vid&"'"

conn.execute (query)

if err.number <> 0 then
	Response.Write("<commit>false</commit>")
else
	Response.Write("<commit>true</commit>")
end if
Response.Write("</modificar>") 
%>