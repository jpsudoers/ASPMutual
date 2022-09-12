<!--#include file="../conexion.asp"-->
<!--#include file="../cnn_string.asp"-->
<%
Response.ContentType = "text/xml"
Response.AddHeader "Cache-control", "private"
Response.AddHeader "Expires", "-1"
Response.CodePage = 65001
Response.CharSet = "utf-8"
%>
<%
Dim DATOS
Dim oConn
SET oConn = Server.CreateObject("ADODB.Connection")
'oConn.Open("Provider=SQLOLEDB; User ID=sa;Password=SCL.2013.2013;data source=.\SQLEXPRESS;Initial Catalog=dbmas")
oConn.Open(MM_cnn_STRING)
Set DATOS = Server.CreateObject("ADODB.RecordSet")
DATOS.CursorType=3

vRutEmpresa = Request("rut_empresa")
vRutUsuario = Request("rut_usuario")
vPassUsuario = Request("pass")
vTipo = Request("tipo")

if(vTipo="0")then
sql = "select dbo.MayMinTexto (R_SOCIAL) as R_SOCIAL,ID_EMPRESA as usuario,TIPO_CONTACTO,ACTIVA from EMPRESAS where RUT='"&vRutEmpresa&"'"
sql = sql&" and EMAIL='"&vRutUsuario&"' and PASSWORD_COORDINACION='"&vPassUsuario&"'"
sql = sql&" or RUT='"&vRutEmpresa&"' and EMAIL_CONTA='"&vRutUsuario&"' and PASSWORD_CONTA='"&vPassUsuario&"'"
else
sql = "select ID_USUARIO as usuario, NOMBRES+' '+A_PATERNO as nombre from USUARIOS "
sql = sql&" where RUT='"&vRutUsuario&"' and CONTRASEÑA='"&vPassUsuario&"'"
end if

DATOS.Open sql,oConn


Response.Write("<?xml version='1.0' encoding='utf-8' ?>"&chr(13))
Response.Write("<DATOS>"&chr(13)) 

     if(DATOS.RecordCount<>"0")then
	    Session.TimeOut = 20
		if(vTipo<>"0")then
		   Session("usuarioMutual") = DATOS("usuario")
		   Session("nombre") = DATOS("nombre")
		else
		   Session("usuario") = DATOS("usuario")
		   Session("nombre") = DATOS("R_SOCIAL")
		   Session("tipo_user_empresa") = DATOS("TIPO_CONTACTO")
		   Session("activa_empresa") = DATOS("ACTIVA")
		   
		   dim cordinador
		   cordinador= " select dbo.MayMinTexto (em.NOMBRES) as NOMBRES from EMPRESAS em "
		   cordinador= cordinador&" where em.EMAIL='"&Request("rut_usuario")&"' and em.PASSWORD_COORDINACION='"&Request("pass")&"'"
		   set rsCordinador = conn.execute (cordinador)

		   if not rsCordinador.eof and not rsCordinador.bof then 
			   Session("correo_user") = rsCordinador("NOMBRES")
			   Session("cargo_user_empresa") = "1"
		   end if
		  
		   dim contabilidad
		   contabilidad= "select dbo.MayMinTexto (em.NOMBRE_CONTA) as NOMBRE_CONTA from EMPRESAS em"
		   contabilidad= contabilidad&" where em.EMAIL_CONTA='"&Request("rut_usuario")&"' and em.PASSWORD_CONTA='"&Request("pass")&"'"
		   set rsContabilidad = conn.execute (contabilidad)

		   if not rsContabilidad.eof and not rsContabilidad.bof then 
			   Session("correo_user") = rsContabilidad("NOMBRE_CONTA")
			   Session("cargo_user_empresa") = "0"
		   end if
		end if
		
		Session("tipoUsuario") = vTipo

		Response.Write("<row>"&chr(13))
		Response.Write("<REGISTRO>"&DATOS.RecordCount&"</REGISTRO>"&chr(13))
		Response.Write("</row>"&chr(13))
	else
		Response.Write("<row>"&chr(13))
		Response.Write("<REGISTRO>"&DATOS.RecordCount&"</REGISTRO>"&chr(13))
		Response.Write("</row>"&chr(13))
	end if
Response.Write("</DATOS>") 
%>
