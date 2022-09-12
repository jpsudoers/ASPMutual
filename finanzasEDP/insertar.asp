<!--#include file="../conexion.asp"-->
<%Response.CodePage = 65001
Response.CharSet = "utf-8"
Response.ContentType="text/xml"
Response.Write("<?xml version='1.0' encoding='UTF-8'?>")
Response.Write("<insertar>") %>
<%
on error resume next
set objUpload = new xelUpload
vempresa = Session("usuario")
vnroOC = objUpload.Form("txtNroOC")
vDescripcion = objUpload.Form("txtDescripcionOC")
vmonto = objUpload.Form("txtMon")
vArchivo = objUpload.Form("txtArchi")

prog=""
if(Request("meses")<>"0" or Request("ano")<>"0")then
	prog=" and "
	if(Request("meses")<>"0")then
		prog=prog&"Month(F.FECHA_EMISION)='"&Request("meses")&"'"
	end if
	
	if(Request("ano")<>"0")then
		if(Request("meses")<>"0")then
			prog=prog&" and "
		end if
	
		prog=prog&" Year(F.FECHA_EMISION)='"&Request("ano")&"'"
	end if

	'prog=" and A.ID_PROGRAMA in (select p2.ID_PROGRAMA from PROGRAMA p2 where "
	'if(Request("selMes")<>"0")then
		'prog=prog&"Month(p2.FECHA_INICIO_)='"&Request("selMes")&"'"
	'end if
	
	'if(Request("selAno")<>"0")then
		'if(Request("selMes")<>"0")then
			'prog=prog&" and "
		'end if
	
		'prog=prog&" Year(p2.FECHA_INICIO_)='"&Request("selAno")&"'"
	'end if
	
	'prog=prog&")"
end if

emp=""
if(Request("Empresa")<>"")then
emp=" and F.ID_EMPRESA='"&Request("Empresa")&"'"
end if

est=""
if(Request("tipo")<>"")then
	est=" and F.ESTADO='"&Request("tipo")&"'"
end if

bdoc=""
if(Request("doc")<>"")then
	bdoc=" and F.N_DOCUMENTO LIKE '%"&Request("doc")&"%'"
end if
set objFich = objUpload.Ficheros("txtArchi")
aux=split(objFich.Nombre,".")
ext=aux(ubound(aux))

fecha = year(now)&right("0"&month(now()),2)&right("0"&day(now()),2)
fecha = fecha&right("0"&hour(now()),2)&right("0"&minute(now()),2)&right("0"&second(now()),2)

nom_arch="Documento_"&fecha&"_"&vempresa&"."&ext


dim preins

preins =" update F"
preins = preins&" set F.DOCUMENTO_COMPROMISO='"&nom_arch&"', F.N_DOCUMENTO='"&vnroOC&"'"
preins = preins&" from FACTURAS F"
preins = preins&" inner join AUTORIZACION A on A.ID_AUTORIZACION=F.ID_AUTORIZACION "
preins = preins&" inner join EMPRESAS EI ON EI.ID_EMPRESA=A.ID_EMPRESA "
preins = preins&" inner join EMPRESAS E on E.ID_EMPRESA=F.ID_EMPRESA "
preins = preins&" inner join PROGRAMA P on P.ID_PROGRAMA=A.ID_PROGRAMA "  
preins = preins&" inner join CURRICULO C on C.ID_MUTUAL=P.ID_MUTUAL " 
preins = preins&" inner join bloque_programacion BQ on BQ.id_bloque=A.ID_BLOQUE " 
preins = preins&" inner join  EMPRESA_CONDICION_COMERCIAL cc on E.ID_EMPRESA = cc.ID_EMPRESA"
preins = preins&" WHERE A.ESTADO in (0,1) and  cc.ID_CONDICION_COMERCIAL in (0,5) and F.FECHA_PAGO is null and F.DESCRIPCION_ESTADO is null "&prog&emp&est&bdoc

Response.Write(preins)
Response.End()
'conn.execute (preins)

dim preinsAUTO

preinsAUTO =" update A"
preinsAUTO = preinsAUTO&" set A.DOCUMENTO_COMPROMISO='"&nom_arch&"'"
preinsAUTO = preinsAUTO&" from AUTORIZACION A"
preinsAUTO = preinsAUTO&" inner join FACTURAS F on A.ID_AUTORIZACION=F.ID_AUTORIZACION "
preinsAUTO = preinsAUTO&" inner join EMPRESAS EI ON EI.ID_EMPRESA=A.ID_EMPRESA "
preinsAUTO = preinsAUTO&" inner join EMPRESAS E on E.ID_EMPRESA=F.ID_EMPRESA "
preinsAUTO = preinsAUTO&" inner join PROGRAMA P on P.ID_PROGRAMA=A.ID_PROGRAMA "  
preinsAUTO = preinsAUTO&" inner join CURRICULO C on C.ID_MUTUAL=P.ID_MUTUAL " 
preinsAUTO = preinsAUTO&" inner join bloque_programacion BQ on BQ.id_bloque=A.ID_BLOQUE " 
preinsAUTO = preinsAUTO&" inner join  EMPRESA_CONDICION_COMERCIAL cc on E.ID_EMPRESA = cc.ID_EMPRESA"
preinsAUTO = preinsAUTO&" WHERE A.ESTADO in (0,1) and  cc.ID_CONDICION_COMERCIAL in (0,5) and F.FECHA_PAGO is null and F.DESCRIPCION_ESTADO is null "&prog&emp&est&bdoc

'conn.execute (preinsAUTO)

next

if err.number <> 0 then
	Response.Write("<commit>false</commit>")
else
	Response.Write("<commit>true</commit>")
end if
Response.Write("</insertar>") 
%>