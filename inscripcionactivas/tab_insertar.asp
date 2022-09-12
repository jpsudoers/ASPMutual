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
vtrabhistid = Request("TrabHistId")
vtrabid = Request("txtTrabID")
vcheckExtran = Request("checkExtran")
vtrabfecha = Request("TrabFecha")
VCORREO = Request("txtMail")

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
	trab = "insert into TRABAJADOR (RUT,NOMBRES,CARGO_EMPRESA,ESCOLARIDAD,ESTADO,NOM_TRAB,APATERTRAB,AMATERTRAB,NACIONALIDAD,"
	trab = trab&"ID_EXTRANJERO,CORREO,EMAIL) "
	trab = trab&" values('"&vruttrab&"',dbo.MayMinTexto('"&vnomtrab&" "&vapatertrab&" "&vamatertrab&"'),"
	trab = trab&"dbo.MayMinTexto('"&vcargotrab&"'),'"&vescolaridadtrab&"',1,"
	trab = trab&"dbo.MayMinTexto('"&vnomtrab&"'),dbo.MayMinTexto('"&vapatertrab&"'),dbo.MayMinTexto('"&vamatertrab&"'),"
	trab = trab&"'"&vcheckExtran&"','"&vIdNacion&"','"&VCORREO&"',"&emailTrab&") "
	conn.execute (trab)
	
	set rsTrab = conn.execute ("select ID_TRABAJADOR from TRABAJADOR where RUT='"&vruttrab&"'")
	trabajador_id=rsTrab("ID_TRABAJADOR")
else
	dim trab_up
	trab_up = "update TRABAJADOR set NOMBRES=dbo.MayMinTexto('"&vnomtrab&" "&vapatertrab&" "&vamatertrab&"'),"
	trab_up = trab_up&"CARGO_EMPRESA=dbo.MayMinTexto('"&vcargotrab&"'),ESCOLARIDAD='"&vescolaridadtrab&"',"
	trab_up = trab_up&"NOM_TRAB=dbo.MayMinTexto('"&vnomtrab&"'),APATERTRAB=dbo.MayMinTexto('"&vapatertrab&"'),"
	trab_up = trab_up&"AMATERTRAB=dbo.MayMinTexto('"&vamatertrab&"'),CORREO='"&VCORREO&"',EMAIL="&emailTrab&" where RUT='"&vruttrab&"' and ID_TRABAJADOR='"&vtrabid&"'"
	conn.execute (trab_up)
	
	set rsRegTrab = conn.execute ("select ID_TRABAJADOR from TRABAJADOR where RUT='"&vruttrab&"'")
	trabajador_id=rsRegTrab("ID_TRABAJADOR")
end if

dim upHistTrab
upHistTrab = "update HISTORICO_CURSOS set TRABIDUP='"&trabajador_id&"', TRABUP='"&vtrabfecha&"'" 
upHistTrab = upHistTrab&" where ID_HISTORICO_CURSO='"&vtrabhistid&"'"

conn.execute (upHistTrab)

if err.number <> 0 then
	Response.Write("<commit>false</commit>")
else
	Response.Write("<commit>true</commit>")
end if
Response.Write("</insertar>") 
%>