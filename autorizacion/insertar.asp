<%Response.CodePage = 65001
Response.CharSet = "utf-8"
Response.ContentType="text/xml"
Response.Write("<?xml version='1.0' encoding='UTF-8'?>")
Response.Write("<insertar>") %>
<!--#include file="../conexion.asp"-->
<%
on error resume next
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

dim query
query = "insert into AUTORIZACION (ID_PROGRAMA,ID_EMPRESA,ID_OTIC,ORDEN_COMPRA,VALOR_OC,CODIGO_AUTORIZACION,FECHA__AUTORIZACION"
query = query&",VALOR_AUTORIZADO,INSCRITOS,ESTADO,ORDEN_COMPRA_OTIC,VALOR_OCOMPRA_OTIC,SOLICITUD)"
query = query&" values('"&vcurriculo&"','"&vempresa&"','"&votic&"','"&vordenemp&"','"&vvalor&"','1232',CONVERT(datetime,'"&vfecha&"',105)"
query = query&",'"&vvalorautorizador&"','"&vinscrito&"',1,'"&vordenotic&"','"&vvalorotic&"','"&vsol&"') "

'RESPONSE.Write(query)
'RESPONSE.End()

Response.Write("<sql>"&query&"</sql>")
conn.execute (query)
if err.number <> 0 then
	Response.Write("<commit>false</commit>")
else
	Response.Write("<commit>true</commit>")
end if
Response.Write("<mensaje>"&vmensaje&"</mensaje>")
Response.Write("</insertar>") 
%>