<!--#include file="../funciones/xelupload.asp"-->
<!--#include file="../conexion.asp"-->
<%
on error resume next
Dim objUpload, objFich

set objUpload = new xelUpload
objUpload.Upload()

'vempresa = Session("usuario")
vnroOC = objUpload.Form("txtNroOC")
vDescripcion = objUpload.Form("txtDescripcionOC")
vmonto = objUpload.Form("txtMon")
vArchivo = objUpload.Form("txtArchi")
meses = objUpload.Form("meses")
ano =  objUpload.Form("ano")
Empresa = objUpload.Form("Empresa")
Estado = objUpload.Form("tipo")
doc = objUpload.Form("doc")
prog=""
prog2=""
if(meses<>"0")then
 prog=" and Month(F.FECHA_EMISION)="&meses
end if
if(ano<>"0")then
	prog2=" and Year(F.FECHA_EMISION)="&ano
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
'end if

emp=""
if(Empresa<>"")then
emp=" and F.ID_EMPRESA="&Empresa
end if

est=""
if(Estado<>"")then
	est=" and F.ESTADO='"&Estado&"'"
end if

bdoc=""
if(doc<>"")then
	bdoc=" and F.N_DOCUMENTO LIKE '%"&doc&"%'"
end if
set objFich = objUpload.Ficheros("txtArchi")
aux=split(objFich.Nombre,".")
ext=aux(ubound(aux))

fecha = year(now)&right("0"&month(now()),2)&right("0"&day(now()),2)
fecha = fecha&right("0"&hour(now()),2)&right("0"&minute(now()),2)&right("0"&second(now()),2)

nom_arch="Documento_"&fecha&"_"&Empresa&"."&ext



dim preins

preins =" update F"
preins = preins&" set F.DOCUMENTO_COMPROMISO='"&nom_arch&"', F.N_DOCUMENTO='"&vnroOC&"', F.Valor_OC="&vmonto
preins = preins&" from FACTURAS F"
preins = preins&" inner join AUTORIZACION A on A.ID_AUTORIZACION=F.ID_AUTORIZACION "
preins = preins&" inner join EMPRESAS EI ON EI.ID_EMPRESA=A.ID_EMPRESA "
preins = preins&" inner join EMPRESAS E on E.ID_EMPRESA=F.ID_EMPRESA "
preins = preins&" inner join PROGRAMA P on P.ID_PROGRAMA=A.ID_PROGRAMA "  
preins = preins&" inner join CURRICULO C on C.ID_MUTUAL=P.ID_MUTUAL " 
preins = preins&" inner join bloque_programacion BQ on BQ.id_bloque=A.ID_BLOQUE " 
preins = preins&" inner join  EMPRESA_CONDICION_COMERCIAL cc on E.ID_EMPRESA = cc.ID_EMPRESA"
preins = preins&" WHERE A.ESTADO in (0,1) and  cc.ID_CONDICION_COMERCIAL in (0,5) and F.FECHA_PAGO is null and F.DESCRIPCION_ESTADO is null "&prog&prog2&emp&est&bdoc


conn.execute (preins)

dim preinsAUTO

preinsAUTO =" update A"
preinsAUTO = preinsAUTO&" set A.DOCUMENTO_COMPROMISO='"&nom_arch&"', A.ORDEN_COMPRA='"&vnroOC&"', A.Valor_OC="&vmonto
preinsAUTO = preinsAUTO&" from AUTORIZACION A"
preinsAUTO = preinsAUTO&" inner join FACTURAS F on A.ID_AUTORIZACION=F.ID_AUTORIZACION "
preinsAUTO = preinsAUTO&" inner join EMPRESAS EI ON EI.ID_EMPRESA=A.ID_EMPRESA "
preinsAUTO = preinsAUTO&" inner join EMPRESAS E on E.ID_EMPRESA=F.ID_EMPRESA "
preinsAUTO = preinsAUTO&" inner join PROGRAMA P on P.ID_PROGRAMA=A.ID_PROGRAMA "  
preinsAUTO = preinsAUTO&" inner join CURRICULO C on C.ID_MUTUAL=P.ID_MUTUAL " 
preinsAUTO = preinsAUTO&" inner join bloque_programacion BQ on BQ.id_bloque=A.ID_BLOQUE " 
preinsAUTO = preinsAUTO&" inner join  EMPRESA_CONDICION_COMERCIAL cc on E.ID_EMPRESA = cc.ID_EMPRESA"
preinsAUTO = preinsAUTO&" WHERE A.ESTADO in (0,1) and  cc.ID_CONDICION_COMERCIAL in (0,5) and F.FECHA_PAGO is null and F.DESCRIPCION_ESTADO is null "&prog&prog2&emp&est&bdoc

conn.execute (preinsAUTO)
objFich.GuardarComo nom_arch,Server.MapPath("../ordenes")






%>