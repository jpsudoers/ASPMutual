<!--#include file="../conexion.asp"-->
<%
Response.CodePage = 65001
Response.CharSet = "utf-8"
Response.ContentType="text/xml"
Response.Write("<?xml version='1.0' encoding='UTF-8'?>")
Response.Write("<modificar>") 

on error resume next
vId  = Request("txtId")
vcurriculo = Request("Curriculo")
vempresa = Request("Empresa")
vordenemp = Request("txtOrdenCompraE")
vvalor=Request("txtValorE")
votic = Request("OTIC")
vordenotic = Request("txtOrdenCompraO")
vvalorotic=Request("txtValorO")
vfecha=Request("txtFecha")
vautorizador=Request("txtAutorizador")
vvalorautorizador=Request("txtValorAutorizador")
vinscrito=Request("txtInscrito")


query="UPDATE AUTORIZACION SET ID_MUTUAL='"&vcurriculo&"',ID_EMPRESA='"&vempresa&"',ID_OTIC='"&votic&"',"
query=query&"ORDEN_COMPRA='"&vordenemp&"',VALOR_OC='"&vvalor&"',"
query=query&"FECHA__AUTORIZACION=CONVERT(datetime,'"&vfecha&"',105),AUTORIZADOR='"&vautorizador&"',"
query=query&"VALOR_AUTORIZADO='"&vvalorautorizador&"',INSCRITOS='"&vvalorautorizador&"',ORDEN_COMPRA_OTIC='"&vordenotic&"',VALOR_OCOMPRA_OTIC='"&vvalorotic&"' "
query=query&" WHERE ID_AUTORIZACION = '"&vId&"'"

'response.Write(query)
'response.End()

conn.execute (query)
Response.Write("<sql>"&query&"</sql>")
if err.number <> 0 then
	Response.Write("<commit>false</commit>")
else
	Response.Write("<commit>true</commit>")
end if
Response.Write("</modificar>") 
%>