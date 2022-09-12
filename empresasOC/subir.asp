<!--#include file="../funciones/xelupload.asp"-->
<!--#include file="../conexion.asp"-->
<%
on error resume next
Dim objUpload, objFich

set objUpload = new xelUpload
objUpload.Upload()

vempresa = Session("usuario")
vnroOC = objUpload.Form("txtNroOC")
vDescripcion = objUpload.Form("txtDescripcionOC")
vmonto = objUpload.Form("txtMon")
vArchivo = objUpload.Form("txtArchi")

set objFich = objUpload.Ficheros("txtArchi")
aux=split(objFich.Nombre,".")
ext=aux(ubound(aux))

fecha = year(now)&right("0"&month(now()),2)&right("0"&day(now()),2)
fecha = fecha&right("0"&hour(now()),2)&right("0"&minute(now()),2)&right("0"&second(now()),2)

nom_arch="Documento_"&fecha&"_"&vempresa&"."&ext


dim preins
preins = "IF NOT EXISTS (select p.* from EMPRESAS_OC p where p.NRO_OC='"&vnroOC&"') BEGIN "
preins = preins&"insert into EMPRESAS_OC(ID_EMPRESA,NRO_OC,MONTOOC,MONTOUTILIZADO,DESCRIPCION,NOMBREARCHIVO,RUTA_ARCHIVO,FECHACREACION)"
preins = preins&" values("&vempresa&",'"&vnroOC&"',"&vmonto&",0,'"&vDescripcion&"','"&nom_arch&"', null ,GETDATE ()) END"



conn.execute (preins)
objFich.GuardarComo nom_arch,Server.MapPath("../ordenes")






%>