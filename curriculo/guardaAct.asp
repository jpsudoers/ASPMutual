<%Response.CodePage = 65001
Response.CharSet = "utf-8"
Response.ContentType="text/xml"
Response.Write("<?xml version='1.0' encoding='UTF-8'?>")
Response.Write("<insertar>") %>
<!--#include file="../conexion.asp"-->
<%
on error resume next
vtxtactId=Request("txtactId")
vactBloque=Request("actBloque")
vactTema=Request("actTema")
vactAct=Request("actAct")
vactHoras=Request("actHoras")
vactDia=Request("actDia")
vtxtIdMutual=Request("txtIdMutual")
vtxtBloqueId=Request("txtBloqueId")


dim query
if(Request("txtactId")="0")then
query = "IF NOT EXISTS (select * from CURRICULO_BLOQUE C where C.ID_MUTUAL='"&vtxtIdMutual&"' and C.ESTADO=1 and c.NOMBRE_BLOQUE='"&vactBloque&"' and c.HORAS='"&vactHoras&"' and c.DIA='"&vactDia&"') BEGIN "
query = query&" insert into CURRICULO_BLOQUE (ID_MUTUAL,NOMBRE_BLOQUE,HORAS,ESTADO,DIA)"
query = query&" values('"&vtxtIdMutual&"','"&vactBloque&"','"&vactHoras&"',1,'"&vactDia&"'); "
query = query&" insert into CURRICULO_ACTIVIDADES (ID_BLOQUE_CURSO,TEMA,ACTIVIDAD,ESTADO)"
query = query&" values((select C.ID_BLOQUE_CURSO from CURRICULO_BLOQUE C where C.ID_MUTUAL='"&vtxtIdMutual&"' and C.ESTADO=1 and c.NOMBRE_BLOQUE='"&vactBloque&"' and c.HORAS='"&vactHoras&"' and c.DIA='"&vactDia&"'),dbo.MayMinTexto('"&vactTema&"'),dbo.MayMinTexto('"&vactAct&"'),1);"
query = query&" END ELSE BEGIN "
query = query&" insert into CURRICULO_ACTIVIDADES (ID_BLOQUE_CURSO,TEMA,ACTIVIDAD,ESTADO)"
query = query&" values((select C.ID_BLOQUE_CURSO from CURRICULO_BLOQUE C where C.ID_MUTUAL='"&vtxtIdMutual&"' and C.ESTADO=1 and c.NOMBRE_BLOQUE='"&vactBloque&"' and c.HORAS='"&vactHoras&"' and c.DIA='"&vactDia&"'),dbo.MayMinTexto('"&vactTema&"'),dbo.MayMinTexto('"&vactAct&"'),1);"
query = query&" END"
else
query = "IF NOT EXISTS (select * from CURRICULO_BLOQUE C where C.ID_MUTUAL='"&vtxtIdMutual&"' and C.ESTADO=1 and c.NOMBRE_BLOQUE='"&vactBloque&"' and c.HORAS='"&vactHoras&"' and c.DIA='"&vactDia&"') BEGIN "
query = query&" insert into CURRICULO_BLOQUE (ID_MUTUAL,NOMBRE_BLOQUE,HORAS,ESTADO,DIA)"
query = query&" values('"&vtxtIdMutual&"','"&vactBloque&"','"&vactHoras&"',1,'"&vactDia&"'); "
query = query&" update CURRICULO_ACTIVIDADES set TEMA=dbo.MayMinTexto('"&vactTema&"'), ACTIVIDAD=dbo.MayMinTexto('"&vactAct&"'), ID_BLOQUE_CURSO=(select C.ID_BLOQUE_CURSO from CURRICULO_BLOQUE C where C.ID_MUTUAL='"&vtxtIdMutual&"' and C.ESTADO=1 and c.NOMBRE_BLOQUE='"&vactBloque&"' and c.HORAS='"&vactHoras&"' and c.DIA='"&vactDia&"') where ID_ACTIVIDAD_CURSO='"&vtxtactId&"';"
query = query&"IF NOT EXISTS (select * from CURRICULO_ACTIVIDADES ac where ac.ID_BLOQUE_CURSO='"&vtxtBloqueId&"' and ac.estado=1) "
query = query&" BEGIN UPDATE CURRICULO_BLOQUE SET ESTADO=0 WHERE ID_BLOQUE_CURSO='"&vtxtBloqueId&"' END"
query = query&" END ELSE BEGIN "
query = query&" update CURRICULO_ACTIVIDADES set TEMA=dbo.MayMinTexto('"&vactTema&"'), ACTIVIDAD=dbo.MayMinTexto('"&vactAct&"'), ID_BLOQUE_CURSO=(select C.ID_BLOQUE_CURSO from CURRICULO_BLOQUE C where C.ID_MUTUAL='"&vtxtIdMutual&"' and C.ESTADO=1 and c.NOMBRE_BLOQUE='"&vactBloque&"' and c.HORAS='"&vactHoras&"' and c.DIA='"&vactDia&"') where ID_ACTIVIDAD_CURSO='"&vtxtactId&"';"
query = query&"IF NOT EXISTS (select * from CURRICULO_ACTIVIDADES ac where ac.ID_BLOQUE_CURSO='"&vtxtBloqueId&"' and ac.estado=1) "
query = query&" BEGIN UPDATE CURRICULO_BLOQUE SET ESTADO=0 WHERE ID_BLOQUE_CURSO='"&vtxtBloqueId&"' END"
query = query&" END"
end if

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