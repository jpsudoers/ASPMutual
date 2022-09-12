<!--#include file="../conexion.asp"-->
<%
Response.CodePage = 65001
Response.CharSet = "utf-8"
Response.ContentType="text/xml"
Response.Write("<?xml version='1.0' encoding='UTF-8'?>")
Response.Write("<modificar>") 

on error resume next
vSepSelIdHst = Request("SepSelIdHst")
vIdAuSep = Request("IdAuSep")

set rsAuto = conn.execute ("select ID_PROGRAMA,ID_EMPRESA,ID_OTIC,ORDEN_COMPRA,CONVERT(VARCHAR(10),FECHA__AUTORIZACION, 105) as FECHA_AUTORIZACION,ESTADO,DOCUMENTO_COMPROMISO,ID_BLOQUE,TIPO_DOC,VALOR_CURSO,CON_OTIC,FACTURADO,CON_FRANQUICIA from AUTORIZACION where ID_AUTORIZACION='"&vIdAuSep&"'")

dim query
query = "IF EXISTS (select * from HISTORICO_CURSOS HC where HC.ID_AUTORIZACION='"&vIdAuSep&"'"
query = query&" and HC.ID_HISTORICO_CURSO in ("&vSepSelIdHst&")) BEGIN "
query = query&"insert into AUTORIZACION (ID_PROGRAMA,ID_EMPRESA,ID_OTIC,ORDEN_COMPRA,FECHA__AUTORIZACION,ESTADO,"
query = query&"DOCUMENTO_COMPROMISO,ID_BLOQUE,TIPO_DOC,VALOR_CURSO,CON_OTIC,FACTURADO,CON_FRANQUICIA)"
query = query&" values('"&rsAuto("ID_PROGRAMA")&"','"&rsAuto("ID_EMPRESA")&"','"&rsAuto("ID_OTIC")&"',"
query = query&"'"&rsAuto("ORDEN_COMPRA")&"',CONVERT(datetime,'"&rsAuto("FECHA_AUTORIZACION")&"',105),1,"
query = query&"'"&rsAuto("DOCUMENTO_COMPROMISO")&"','"&rsAuto("ID_BLOQUE")&"',4,'"&rsAuto("VALOR_CURSO")&"',0,1,0) END"

conn.execute (query)

set rsUltAuto = conn.execute ("select IDENT_CURRENT('AUTORIZACION')AS UltAuto")

dim ultRegAuto
ultRegAuto=rsUltAuto("UltAuto")

conn.execute ("UPDATE HISTORICO_CURSOS set ID_AUTORIZACION="&ultRegAuto&" where ID_HISTORICO_CURSO in ("&vSepSelIdHst&") and ID_AUTORIZACION="&vIdAuSep)

dim upAuto
upAuto="UPDATE AUTORIZACION SET N_PARTICIPANTES=(SELECT COUNT(*) FROM HISTORICO_CURSOS WHERE ID_AUTORIZACION="&vIdAuSep&"), "
upAuto=upAuto&"AUTORIZACION.VALOR_OC=AUTORIZACION.VALOR_CURSO*(SELECT COUNT(*) FROM HISTORICO_CURSOS " 
upAuto=upAuto&"WHERE ID_AUTORIZACION="&vIdAuSep&") WHERE ID_AUTORIZACION="&vIdAuSep

conn.execute (upAuto)

dim upAutoSeg
upAutoSeg="UPDATE AUTORIZACION SET N_PARTICIPANTES=(SELECT COUNT(*) FROM HISTORICO_CURSOS WHERE ID_AUTORIZACION="&ultRegAuto&"), "
upAutoSeg=upAutoSeg&"AUTORIZACION.VALOR_OC=AUTORIZACION.VALOR_CURSO*(SELECT COUNT(*) FROM HISTORICO_CURSOS " 
upAutoSeg=upAutoSeg&"WHERE ID_AUTORIZACION="&ultRegAuto&") WHERE ID_AUTORIZACION="&ultRegAuto

conn.execute (upAutoSeg)

'query=vSepSelIdHst&" "&vIdAuSep
Response.Write("<sql>"&query&"</sql>")
if err.number <> 0 then
	Response.Write("<commit>false</commit>")
else
	Response.Write("<commit>true</commit>")
end if
Response.Write("</modificar>") 
%>