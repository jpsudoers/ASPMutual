<!--#include file="../conexion.asp"-->
<%
Response.CodePage = 65001
Response.CharSet = "utf-8"
Response.ContentType="text/xml"
Response.Write("<?xml version='1.0' encoding='UTF-8'?>")
Response.Write("<modificar>") 

on error resume next
vUSSelIdAu = Request("USSelIdAu")
vIdAuMan = Request("IdAuMan")
vUSIdbloque = Request("USIdbloque")
vUSIdRelator = Request("USIdRelator")
vUSIdSede = Request("USIdSede")

dim upHist
upHist="UPDATE HISTORICO_CURSOS set ID_AUTORIZACION="&vIdAuMan&", RELATOR="&vUSIdRelator&","
upHist=upHist&" SEDE="&vUSIdSede&", ID_BLOQUE="&vUSIdbloque&" where ID_AUTORIZACION in ("&vUSSelIdAu&")"

conn.execute (upHist)

conn.execute ("delete from AUTORIZACION where AUTORIZACION.ID_AUTORIZACION in ("&vUSSelIdAu&")")

dim upAuto
upAuto="UPDATE AUTORIZACION SET N_PARTICIPANTES=(SELECT COUNT(*) FROM HISTORICO_CURSOS WHERE ID_AUTORIZACION="&vIdAuMan&"), "
upAuto=upAuto&"AUTORIZACION.VALOR_OC=AUTORIZACION.VALOR_CURSO*(SELECT COUNT(*) FROM HISTORICO_CURSOS " 
upAuto=upAuto&"WHERE ID_AUTORIZACION="&vIdAuMan&") WHERE ID_AUTORIZACION="&vIdAuMan

conn.execute (upAuto)

'query="UPDATE HISTORICO_CURSOS set ID_AUTORIZACION="&vIdAuMan&" where ID_AUTORIZACION in ("&vUSSelIdAu&") "&"delete from AUTORIZACION where AUTORIZACION.ID_AUTORIZACION in ("&vUSSelIdAu&") "&upAuto

Response.Write("<sql>"&upHist&"</sql>")
if err.number <> 0 then
	Response.Write("<commit>false</commit>")
else
	Response.Write("<commit>true</commit>")
end if
Response.Write("</modificar>") 
%>