<%Response.CodePage = 65001
Response.CharSet = "utf-8"
Response.ContentType="text/xml"
Response.Write("<?xml version='1.0' encoding='UTF-8'?>")
Response.Write("<insertar>") %>
<!--#include file="../conexion.asp"-->
<%
on error resume next
vruttrab = Request("txtRutTrab")
vcargotrab = trim(Request("txtCargoTrab"))
vescolaridadtrab = Request("escolaridadTrab")
vnomtrab = trim(Request("txtNomTrab"))
vapatertrab = trim(Request("txtAPaterTrab"))
vamatertrab = trim(Request("txtAMaterTrab"))
vtrabpreins = Request("txtTrabPreins")
vtrabid = Request("txtTrabID")
vcheckExtran = Request("checkExtran")
VCORREO = Request("txtMail")

vGerenciatrab = "Null"
if(trim(Request("txtGerencia"))<>"")then
	vGerenciatrab = "dbo.MayMinTexto('"&trim(Request("txtGerencia"))&"')"
end if

vNSaptrab = "Null"
if(Request("txtNSap")<>"")then
	vNSaptrab = "'"&Request("txtNSap")&"'"
end if

emailTrab = "Null"
if(Request("txtEmailTrab")<>"")then
	emailTrab = "'"&Request("txtEmailTrab")&"'"
end if

if(vtrabid="0")then

	vIdNacion=""
	if(vcheckExtran="1")then
		vIdNacion=Request("txtPasTrab")
	end if

	dim trab
	trab = "IF NOT EXISTS (select * from trabajador t where t.RUT='"&vruttrab&"') BEGIN "
	trab = trab&"insert into TRABAJADOR (RUT,NOMBRES,CARGO_EMPRESA,ESCOLARIDAD,ESTADO,NOM_TRAB,APATERTRAB,AMATERTRAB,NACIONALIDAD,"
	trab = trab&"ID_EXTRANJERO,GERENCIA,NSAP,CORREO,EMAIL) "
	trab = trab&" values('"&vruttrab&"',dbo.MayMinTexto('"&vnomtrab&" "&vapatertrab&" "&vamatertrab&"'),"
	trab = trab&"dbo.MayMinTexto('"&vcargotrab&"'),'"&vescolaridadtrab&"',1,"
	trab = trab&"dbo.MayMinTexto('"&vnomtrab&"'),dbo.MayMinTexto('"&vapatertrab&"'),dbo.MayMinTexto('"&vamatertrab&"'),"
	trab = trab&"'"&vcheckExtran&"','"&vIdNacion&"',"&vGerenciatrab&","&vNsaptrab&",'"&VCORREO&"',"&emailTrab&") END"
	conn.execute (trab)
	
	conn.execute ("insert into Log_Update_Insert values("&Session("usuarioMutual")&",getdate(),'"&Replace(trab,"'","`")&"','"&Request.ServerVariables("PATH_INFO")&"','Inscripcion de Cursos');")
	
	set rsTrab = conn.execute ("select ID_TRABAJADOR from TRABAJADOR where RUT='"&vruttrab&"'")
	trabajador_id=rsTrab("ID_TRABAJADOR")
else
	dim trab_up
	trab_up = "update TRABAJADOR set NOMBRES=dbo.MayMinTexto('"&vnomtrab&" "&vapatertrab&" "&vamatertrab&"'),"
	trab_up = trab_up&"CARGO_EMPRESA=dbo.MayMinTexto('"&vcargotrab&"'),ESCOLARIDAD='"&vescolaridadtrab&"',"
	trab_up = trab_up&"NOM_TRAB=dbo.MayMinTexto('"&vnomtrab&"'),APATERTRAB=dbo.MayMinTexto('"&vapatertrab&"'),"
	trab_up = trab_up&"AMATERTRAB=dbo.MayMinTexto('"&vamatertrab&"'),GERENCIA="&vGerenciatrab&", NSAP="&vNSaptrab&",CORREO='"&VCORREO&"',EMAIL="&emailTrab&" where RUT='"&vruttrab&"' and ID_TRABAJADOR='"&vtrabid&"'"
	conn.execute (trab_up)
	
	conn.execute ("insert into Log_Update_Insert values("&Session("usuarioMutual")&",getdate(),'"&Replace(trab_up,"'","`")&"','"&Request.ServerVariables("PATH_INFO")&"','Inscripcion de Cursos');")
	
	set rsRegTrab = conn.execute ("select ID_TRABAJADOR from TRABAJADOR where RUT='"&vruttrab&"'")
	trabajador_id=rsRegTrab("ID_TRABAJADOR")
end if

dim preIns_Trab
	preIns_Trab = "IF NOT EXISTS (select * from PREINSCRIPCION_TRABAJADOR p where p.id_trabajador='"&trabajador_id&"'"
	preIns_Trab = preIns_Trab&" and p.preinscripcionTemp='"&vtrabpreins&"') BEGIN "	
	preIns_Trab = preIns_Trab&"insert into PREINSCRIPCION_TRABAJADOR(id_trabajador,franquicia,preinscripcionTemp) "
	preIns_Trab = preIns_Trab&" values('"&trabajador_id&"',100,'"&vtrabpreins&"') END"
conn.execute (preIns_Trab)

conn.execute ("insert into Log_Update_Insert values("&Session("usuarioMutual")&",getdate(),'"&Replace(preIns_Trab,"'","`")&"','"&Request.ServerVariables("PATH_INFO")&"','Inscripcion de Cursos');")

conn.execute ("exec dbo.setEscolaridadTrabajadores;")
on error resume next
if err.number <> 0 then
	Response.Write("<commit>false</commit>")
else
	Response.Write("<commit>true</commit>")
end if
Response.Write("</insertar>") 
%>