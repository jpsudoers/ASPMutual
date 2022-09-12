<!--#include file="../funciones/xelupload.asp"-->
<!--#include file="../conexion.asp"-->
<%
on error resume next
Dim objUpload, objFich

set objUpload = new xelUpload
objUpload.Upload()

set objFich = objUpload.Ficheros("txtDoc")

aux=split(objFich.Nombre,".")
ext=aux(ubound(aux))

fecha = year(now)&right("0"&month(now()),2)&right("0"&day(now()),2)
fecha = fecha&right("0"&hour(now()),2)&right("0"&minute(now()),2)&right("0"&second(now()),2)

nom_arch="Documento_"&fecha&".pdf"&ext

objFich.GuardarComo nom_arch,Server.MapPath("../ordenes")

set oFich = nothing
set objUpload = nothing

%>