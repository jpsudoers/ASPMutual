<!--#include file="../conexion.asp"-->
<%
Response.CodePage = 65001
Response.CharSet = "utf-8"
Response.ContentType="text/xml"
Response.Write("<?xml version='1.0' encoding='UTF-8'?>")
Response.Write("<modificar>") 

on error resume next

dim totalMan
totalMan = cdbl(Request("countFilas"))

dim estado_eval

For i = 1 To Request("countFilas") Step 1
if(Request("E"&i)="A")then
	estado_eval="Aprobado"
end if

if(Request("E"&i)="R")then
	estado_eval="Reprobado"
end if

if(Request("E"&i)="C")then
	estado_eval="Con Obs."
end if

query = "update HISTORICO_CURSOS set ASISTENCIA='"&Request("A"&i)&"', "
query = query&"CALIFICACION='"&Request("C"&i)&"', EVALUACION='"&estado_eval&"', ESTADO=2 "
query = query&" where ID_HISTORICO_CURSO='"&Request("H"&i)&"'"

if(Request("A"&i)="0")then
	totalMan = totalMan - 1
end if

conn.execute (query)

Next

set rsART = conn.execute ("select CA.ID_ARTICULO from CURRICULO_ARTICULOS CA where CA.ID_MUTUAL='"&Request("txtIdMutual")&"'")

dim mov
dim upArt
while not rsART.eof
mov = "insert into MOVIMIENTOS (ID_ARTICULO,TIPO_MOVIMIENTO,CANTIDAD,ESTADO,FECHA,ID_BODEGA,MODULO,ID_PROG_BLOQUE,ID_USUARIO)"
mov = mov&" values('"&rsART("ID_ARTICULO")&"',3,'"&"-"&totalMan&"',1,GETDATE (),"
mov = mov&"1,2,(SELECT HC.ID_BLOQUE FROM HISTORICO_CURSOS HC where HC.ID_HISTORICO_CURSO='"&Request("H1")&"'),'"&Session("usuarioMutual")&"') "

conn.execute (mov)


upArt = "update ARTICULO_BODEGA set FECHA_ULT_SALIDA=GETDATE(),STOCK_ACTUAL="
upArt = upArt&"(select STOCK_ACTUAL from ARTICULO_BODEGA where ESTADO_ARTICULO_BODEGA=1 "
upArt = upArt&" AND ID_ARTICULO IS NOT NULL AND ID_ARTICULO='"&rsART("ID_ARTICULO")&"' and ID_BODEGA=1)-"&totalMan
upArt = upArt&" WHERE (ESTADO_ARTICULO_BODEGA = 1) AND (ID_ARTICULO IS NOT NULL) "
upArt = upArt&" and ID_ARTICULO='"&rsART("ID_ARTICULO")&"' and ID_BODEGA=1"

conn.execute (upArt)

rsART.Movenext
wend

Response.Write("<sql>"&query&"</sql>")
if err.number <> 0 then
	Response.Write("<commit>false</commit>")
else
	Response.Write("<commit>true</commit>")
end if
Response.Write("</modificar>") 
%>