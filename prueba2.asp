<!--#include file="conexion.asp"-->
<!--#include file="cnn_string.asp"-->
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

vRutEmpresa = ""
vRutUsuario = "16313889-1"
vPassUsuario = "Mut$2021"
vTipo = "1"

if(vTipo="0")then
sql = "select dbo.MayMinTexto(R_SOCIAL) R_SOCIAL,ID_EMPRESA usuario,TIPO_CONTACTO,ACTIVA,"
sql = sql&"blq=ISNULL((SELECT ETC.ESTADO_EMPRESA_COMPROMISO FROM EMPRESA_TIPO_COMPROMISO ETC WHERE ETC.ID_EMPRESA=EMPRESAS.ID_EMPRESA AND ETC.ID_COMPROMISO_PAGO=1),1),infoBanner=(select count(*) from BANNER_EMPRESAS e where e.ID_EMPRESA=EMPRESAS.ID_EMPRESA and convert(date, e.FECHA_VIGENCIA)>=CONVERT(date, GETDATE())) from EMPRESAS "
sql = sql&" where RUT='"&vRutEmpresa&"' and EMAIL='"&vRutUsuario&"' and PASSWORD_COORDINACION='"&vPassUsuario&"'"
sql = sql&" and ESTADO=1 "
sql = sql&" or RUT='"&vRutEmpresa&"' and EMAIL_CONTA='"&vRutUsuario&"' and PASSWORD_CONTA='"&vPassUsuario&"'"
sql = sql&" and ESTADO=1 "
else
sql = "select ID_USUARIO as usuario, NOMBRES+' '+A_PATERNO as nombre,TIPO_USUARIO as userTipoMutual,"
sql = sql&"ID_INSTRUCTOR as relUser,UNV=ISNULL(PERMISO_NV,0) from USUARIOS "
sql = sql&" where ESTADO=1 and RUT='"&vRutUsuario&"' and CONTRASEÑA='"&vPassUsuario&"'"
end if

DATOS.Open sql,oConn


Response.Write("<?xml version='1.0' encoding='utf-8' ?>"&chr(13))
Response.Write("<DATOS>"&chr(13)) 

     if(DATOS.RecordCount<>"0")then
	    Session.TimeOut = 45
		if(vTipo<>"0")then
		   Session("usuarioMutual") = DATOS("usuario")
		   Session("nombre") = DATOS("nombre")
		   Session("userTipo") = DATOS("userTipoMutual")
		   Session("relUsuario") = DATOS("relUser")
		   Session("U_NV") = DATOS("UNV")
		else
		   Session("usuario") = DATOS("usuario")
		   Session("nombre") = DATOS("R_SOCIAL")
		   Session("tipo_user_empresa") = DATOS("TIPO_CONTACTO")
		   Session("activa_empresa") = DATOS("ACTIVA")
		   Session("countBlq") = DATOS("blq")
	           Session("countInfoB") = DATOS("infoBanner")
		   
		   dim cordinador
		   cordinador= " select dbo.MayMinTexto (em.NOMBRES) as NOMBRES from EMPRESAS em "
		   cordinador= cordinador&" where em.EMAIL='"&Request("rut_usuario")&"' and em.PASSWORD_COORDINACION='"&Request("pass")&"'"
		   set rsCordinador = conn.execute (cordinador)

		   if not rsCordinador.eof and not rsCordinador.bof then 
			   Session("correo_user") = rsCordinador("NOMBRES")
			   Session("cargo_user_empresa") = "1"
		   end if
		  
		   conn.execute ("exec [dbo].[ACTUALIZA_INSCRIPCIONES_AUTORIZADAS] 1")

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
