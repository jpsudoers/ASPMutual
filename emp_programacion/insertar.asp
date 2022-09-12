<%Response.CodePage = 65001
Response.CharSet = "utf-8"
Response.ContentType="text/xml"
Response.Write("<?xml version='1.0' encoding='UTF-8'?>")
Response.Write("<insertar>") %>
<!--#include file="../conexion.asp"-->
<%
on error resume next
vempresa = Request("Empresa")
vprograma = Request("programa")
vempman = Request("EmpMan")
vcompromiso = Request("Compromiso")

vnum = Request("txtNum")

vdoc = Request("txtDoc")
vOrdenOC = Request("selectOC")
vValor = Request("ValorTotal")

if (vnum = "") then
  vnum = vOrdenOC
end



dim preins
preins = "insert into PREINSCRIPCIONES(id_programacion,id_empresa,id_otic,ESTADO,tipo_compromiso,numero_compromiso,"
preins = preins&"doc_compromiso,fecha_preinscripcion) "
preins = preins&" values('"&vprograma&"','"&vempresa&"','"&vempman&"',1,'"&vcompromiso&"','"&vnum&"','"&vdoc&"',GETDATE ()) "
conn.execute (preins)

conn.execute ("insert into Log_Update_Insert values("&Session("usuarioMutual")&",getdate(),'"&Replace(preins,"'","`")&"','"&Request.ServerVariables("PATH_INFO")&"','Inscripcion de Cursos');")

set rs = conn.execute ("select IDENT_CURRENT('PREINSCRIPCIONES')AS UltPreIns")

cadena = split(Request("datostabla"), "/")
trabajador_id=""
for i = 0 to ubound(cadena) - 1

cadena2 = split(cadena(i), ",")

set rsReg = conn.execute ("select COUNT(rut)as total from TRABAJADOR where RUT="&cadena2(0))

if(rsReg("total")="0")then
	dim trab
	trab = "insert into TRABAJADOR (RUT, NOMBRES, CARGO_EMPRESA,ESCOLARIDAD, ESTADO) values("&cadena(i)&",1) "
	conn.execute (trab)
	
	conn.execute ("insert into Log_Update_Insert values("&Session("usuarioMutual")&",getdate(),'"&Replace(trab,"'","`")&"','"&Request.ServerVariables("PATH_INFO")&"','Inscripcion de Cursos');")
	
	set rsTrab = conn.execute ("select IDENT_CURRENT('TRABAJADOR')AS UltTrab")
	trabajador_id=rsTrab("UltTrab")
else
	set rsRegTrab = conn.execute ("select ID_TRABAJADOR from TRABAJADOR where RUT="&cadena2(0))
	trabajador_id=rsRegTrab("ID_TRABAJADOR")
end if

dim preIns_Trab
preIns_Trab = "insert into PREINSCRIPCION_TRABAJADOR(id_preinscripcion,id_trabajador) "
preIns_Trab = preIns_Trab&"values('"&rs("UltPreIns")&"','"&trabajador_id&"')"
conn.execute (preIns_Trab)

conn.execute ("insert into Log_Update_Insert values("&Session("usuarioMutual")&",getdate(),'"&Replace(preIns_Trab,"'","`")&"','"&Request.ServerVariables("PATH_INFO")&"','Inscripcion de Cursos');")

dim empresaoc
empresaoc = "update EMPRESAS_OC"
empresaoc = empresaoc&" set MONTOUTILIZADO="&vValor
empresaoc= empresaoc&"where ordenCompra ="&vOrdenOC&" and id_empresa="&Session("usuario")

conn.execute(empresaoc)

next

if err.number <> 0 then
	Response.Write("<commit>false</commit>")
else
	Response.Write("<commit>true</commit>")
end if
Response.Write("</insertar>") 
%>