<%Response.CodePage = 65001
Response.CharSet = "utf-8"
Response.ContentType="text/xml"
Response.Write("<?xml version='1.0' encoding='UTF-8'?>")
Response.Write("<insertar>") %>
<!--#include file="../conexion.asp"-->
<%
on error resume next
vrut = Request("txRut")
vnomb=trim(Request("txtNomb"))
vap_pater=trim(Request("txtAp_pater"))
vap_mater=trim(Request("txtAp_mater"))
vcel="Null"
if(Request("txtCel")<>"")then
vcel=trim(replace(Request("txtCel")," ",""))
end if
vemail=Request("txtEmail")
vcargo=trim(Request("txtCargo"))
vp1=Request("p1")
vp2=Request("p2")
vp3=Request("p3")
vp4=Request("p4")
vp5=Request("p5")
vp6=Request("p6")
vp7=Request("p7")
vp8=Request("p8")
vp9=Request("p9")
vp10=Request("p10")
vp11=Request("p11")
vp12=Request("p12")
vp13=Request("p13")
vp14=Request("p14")
vp15=Request("p15")
vpass=Request("txtPass")
addRelator=Request("addRelator")

dim addRel

tipUser="0"
idRel="Null"
	if(addRelator="1")then
	
		addRel = "IF NOT EXISTS (select * from INSTRUCTOR_RELATOR ir where ir.RUT='"&vrut&"' AND ir.ESTADO=1) BEGIN "&_
				 " insert into INSTRUCTOR_RELATOR (RUT, A_PATERNO, A_MATERNO, NOMBRES, ESTADO) "&_
				 " values('"&vrut&"',dbo.MayMinTexto('"&vap_pater&"'),dbo.MayMinTexto('"&vap_mater&"'),dbo.MayMinTexto('"&vnomb&"'),1) "&_
				 " select ir.ID_INSTRUCTOR from INSTRUCTOR_RELATOR ir where ir.RUT='"&vrut&"' AND ir.ESTADO=1 end "&_
				 " ELSE begin select ir.ID_INSTRUCTOR from INSTRUCTOR_RELATOR ir where ir.RUT='"&vrut&"' AND ir.ESTADO=1 END "
		
		strSQL = "Set Nocount on "
		strSQL = strSQL + addRel
		strSQL = strSQL + " "
		strSQL = strSQL + " set nocount off"
		
		Set rsRelator = conn.execute (strSQL)
		
		tipUser="1"
		idRel=rsRelator("ID_INSTRUCTOR")
	end if

dim query
query = "IF NOT EXISTS (select * from USUARIOS U where U.rut='"&vrut&"' and U.ESTADO=1) BEGIN "
query = query&"insert into USUARIOS (RUT, A_PATERNO, A_MATERNO, NOMBRES, CELULAR, EMAIL, CARGO_EMPRESA, ESTADO, "
query=query&"PERMISO1, PERMISO2, PERMISO3, PERMISO4,PERMISO5,PERMISO6,PERMISO7,"
query=query&"PERMISO8, PERMISO9, PERMISO10, PERMISO11,PERMISO12,PERMISO13,PERMISO14,PERMISO15,CONTRASEÑA,TIPO_USUARIO,ID_INSTRUCTOR) "
query=query&" values('"&vrut&"',dbo.MayMinTexto('"&vap_pater&"'),dbo.MayMinTexto('"&vap_mater&"'),dbo.MayMinTexto('"&vnomb&"'),"&vcel&","
query=query&"LOWER('"&vemail&"'),dbo.MayMinTexto('"&vcargo&"'),1,'"&vp1&"','"&vp2&"','"&vp3&"','"&vp4&"','"&vp5&"','"&vp6&"','"&vp7&"',"
query=query&"'"&vp8&"','"&vp9&"','"&vp10&"','"&vp11&"','"&vp12&"','"&vp13&"','"&vp14&"','"&vp15&"','"&vpass&"',"&tipUser&","&idRel&") END"

conn.execute (query)

Response.Write("<sql>"&query&"</sql>")
if err.number <> 0 then
	Response.Write("<commit>false</commit>")
else
	Response.Write("<commit>true</commit>")
end if
Response.Write("</insertar>") 
%>