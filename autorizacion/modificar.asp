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
votic = Request("txtIdOtic")
vordenotic = Request("txtOrdenCompraO")
vvalorotic=Request("txtValorO")
vfecha=Request("txtFecha")
vvalorautorizador=Request("txtValorAutorizador")
vinscrito=Request("txtInscrito")
vsol=Request("solicitud")

query="UPDATE AUTORIZACION SET ID_PROGRAMA='"&vcurriculo&"',ID_EMPRESA='"&vempresa&"',ID_OTIC='"&votic&"',"
query=query&"ORDEN_COMPRA='"&vordenemp&"',VALOR_OC='"&vvalor&"',"
query=query&"FECHA__AUTORIZACION=CONVERT(datetime,'"&vfecha&"',105),"
query=query&"VALOR_AUTORIZADO='"&vvalorautorizador&"',INSCRITOS='"&vvalorautorizador&"',ORDEN_COMPRA_OTIC='"&vordenotic&"',"
query=query&"VALOR_OCOMPRA_OTIC='"&vvalorotic&"',SOLICITUD='"&vsol&"' "
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