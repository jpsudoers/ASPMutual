<!--#include file="../conexion.asp"-->
<%Response.CodePage = 65001
Response.CharSet = "utf-8"
Response.ContentType="text/xml"
Response.Write("<?xml version='1.0' encoding='UTF-8'?>")
Response.Write("<insertar>") %>
<%
on error resume next
set objUpload = new xelUpload
objUpload.Upload()
'NroOC;DescripcionOC;txtDir;txtArchi
vempresa = Session("usuario")
vnroOC = Request("txtNroOC")
vDescripcion = Request("txtDescripcionOC")
vmonto = Request"txtMon")
vArchivo = Request("txtArchi")
'archivo
set objFich = objUpload.Ficheros("txtArchi")
aux=split(objFich.Nombre,".")
ext=aux(ubound(aux))

fecha = year(now)&right("0"&month(now()),2)&right("0"&day(now()),2)
fecha = fecha&right("0"&hour(now()),2)&right("0"&minute(now()),2)&right("0"&second(now()),2)

nom_arch="Documento_"&fecha&"_"&vempresa&"."&ext


dim preins
preins = "IF NOT EXISTS (select p.* from EMPRESAS_OC p where p.NRO_OC='"&vnroOC&"') BEGIN "
preins = preins&"insert into EMPRESAS_OC(ID_EMPRESA,NRO_OC,MONTOOC,MONTOUTILIZADO,DESCRIPCION,NOMBREARCHIVO,RUTA_ARCHIVO,FECHACREACION)"
preins = preins&" values("&vempresa&",'"&vnroOC&"',"&vmonto&",0,'"&vDescripcion&"','"&nom_arch&"',GETDATE ()) END"

next

if err.number <> 0 then
	Response.Write("<commit>false</commit>")
else
	Response.Write("<commit>true</commit>")
end if
Response.Write("</insertar>") 
%>