<!--#include file="xelupload.asp"-->
<!--#include file="../conexion.asp"-->
<%
on error resume next
Dim objUpload, objFich

set objUpload = new xelUpload
objUpload.Upload()

vempresa = objUpload.Form("Empresa")
vprograma = objUpload.Form("programa")
vempman = objUpload.Form("EmpMan")
vcompromiso = objUpload.Form("Compromiso")
vnum = objUpload.Form("txtNum")
vparticipantes = objUpload.Form("NParticipantes")
vvalor = objUpload.Form("Valor")
vvalortotal = objUpload.Form("ValorTotal")

vcosto = objUpload.Form("Costo")
vtipoempresa = objUpload.Form("TipoEmpresa")

set objFich = objUpload.Ficheros("txtDoc")

aux=split(objFich.Nombre,".")
ext=aux(ubound(aux))

fecha = year(now)&right("0"&month(now()),2)&right("0"&day(now()),2)
fecha = fecha&right("0"&hour(now()),2)&right("0"&minute(now()),2)&right("0"&second(now()),2)

nom_arch="Documento_"&fecha&"."&ext

dim preins
preins = "insert into PREINSCRIPCIONES(id_programacion,id_empresa,id_otic,ESTADO,tipo_compromiso,numero_compromiso,"
preins = preins&"doc_compromiso,fecha_preinscripcion,participantes,valor,valor_total,costo,tipo_empresa) "
preins = preins&" values('"&vprograma&"','"&vempresa&"','"&vempman&"',1,'"&vcompromiso&"','"&vnum&"','"&nom_arch&"',GETDATE (),"
preins = preins&"'"&vparticipantes&"','"&vvalor&"','"&vvalortotal&"','"&vcosto&"','"&vtipoempresa&"') "
conn.execute (preins)

set rs = conn.execute ("select IDENT_CURRENT('PREINSCRIPCIONES')AS UltPreIns")

cadena = split(objUpload.Form("datostabla"), "/")
porcentaje = split(objUpload.Form("datosfran"), "/")

trabajador_id=""
for i = 0 to ubound(cadena) - 1

cadena2 = split(cadena(i), ",")

set rsReg = conn.execute ("select COUNT(rut)as total from TRABAJADOR where RUT="&cadena2(0))

if(rsReg("total")="0")then
	dim trab
	trab = "insert into TRABAJADOR (RUT, NOMBRES, CARGO_EMPRESA,ESCOLARIDAD, ESTADO) values("&cadena(i)&",1) "
	conn.execute (trab)
	set rsTrab = conn.execute ("select IDENT_CURRENT('TRABAJADOR')AS UltTrab")
	trabajador_id=rsTrab("UltTrab")
else
	set rsRegTrab = conn.execute ("select ID_TRABAJADOR from TRABAJADOR where RUT="&cadena2(0))
	trabajador_id=rsRegTrab("ID_TRABAJADOR")
end if

dim preIns_Trab
preIns_Trab = "insert into PREINSCRIPCION_TRABAJADOR(id_preinscripcion,id_trabajador,franquicia) "
preIns_Trab = preIns_Trab&"values('"&rs("UltPreIns")&"','"&trabajador_id&"',"&porcentaje(i)&")"
conn.execute (preIns_Trab)

next

objFich.GuardarComo nom_arch,Server.MapPath("../ordenes")

set oFich = nothing
set objUpload = nothing
%>