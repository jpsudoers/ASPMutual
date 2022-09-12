<!--#include file="../conexion.asp"-->
<%
Response.CodePage = 65001
Response.CharSet = "utf-8"
Response.ContentType="text/xml"
Response.Write("<?xml version='1.0' encoding='UTF-8'?>")
Response.Write("<modificar>") 

on error resume next
vId  = Request("txtId")
vrut = Request("txtRut")
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

query = "UPDATE USUARIOS SET A_PATERNO=dbo.MayMinTexto('"&vap_pater&"'), A_MATERNO=dbo.MayMinTexto('"&vap_mater&"'),"
query = query&"NOMBRES=dbo.MayMinTexto('"&vnomb&"'), CELULAR="&vcel&", "
query = query&"EMAIL=LOWER('"&vemail&"'), CARGO_EMPRESA=dbo.MayMinTexto('"&vcargo&"'), PERMISO1='"&vp1&"',"
query = query&"PERMISO2='"&vp2&"', PERMISO3='"&vp3&"', PERMISO4='"&vp4&"', PERMISO5='"&vp5&"', "
query = query&"PERMISO6='"&vp6&"', PERMISO7='"&vp7&"', PERMISO8='"&vp8&"', PERMISO9='"&vp9&"', PERMISO10='"&vp10&"', "
query = query&"PERMISO11='"&vp11&"', PERMISO12='"&vp12&"', PERMISO13='"&vp13&"', PERMISO14='"&vp14&"', PERMISO15='"&vp15&"', " 
query = query&"CONTRASEÑA='"&vpass&"',TIPO_USUARIO="&tipUser&",ID_INSTRUCTOR="&idRel&" WHERE ID_USUARIO = '"&vid&"'"

'response.Write(query)
'response.End()

conn.execute (query)
Response.Write("<sql>"&query&"</sql>")
if err.number <> 0 then
	Response.Write("<commit>false</commit>")
else
	Response.Write("<commit>true</commit>")
end if
Response.Write("</modificar>") 
%>