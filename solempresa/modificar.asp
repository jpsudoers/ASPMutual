<!--#include file="../conexion.asp"-->
<%
Response.CodePage = 65001
Response.CharSet = "utf-8"
Response.ContentType="text/xml"
Response.Write("<?xml version='1.0' encoding='UTF-8'?>")
Response.Write("<modificar>") 

Dim my_num,max,min
min=10001
max=50001

Randomize
my_num=int((max-min+1)*rnd+min)

Dim my_num2,max2,min2
min2=10001
max2=50001

Randomize
my_num2=int((max2-min2+1)*rnd+min2)

on error resume next
vId  = Request("txtId")

if(Request("rechazo")="0")then
	if(Request("tipoContacto")="1")then
	query = "UPDATE EMPRESAS SET PREINSCRITA=1,PASSWORD_COORDINACION='"&my_num&"',PASSWORD_CONTA='' "
	query = query&" WHERE id_empresa = '"&vid&"'"
	else
	query = "UPDATE EMPRESAS SET PREINSCRITA=1,PASSWORD_COORDINACION='"&my_num&"',PASSWORD_CONTA='"&my_num2&"' "
	query = query&" WHERE id_empresa = '"&vid&"'"
	end if
	
	conn.execute (query)
	
	queryBloqueo="IF EXISTS (SELECT * FROM EMPRESA_TIPO_COMPROMISO WHERE ID_EMPRESA="&vid&" AND ID_COMPROMISO_PAGO=1) "&_
	      "UPDATE EMPRESA_TIPO_COMPROMISO SET ESTADO_EMPRESA_COMPROMISO=0 WHERE ID_EMPRESA='"&vid&"' AND ID_COMPROMISO_PAGO=1 "&_
	      "ELSE INSERT INTO EMPRESA_TIPO_COMPROMISO(ID_EMPRESA, ID_COMPROMISO_PAGO, ESTADO_EMPRESA_COMPROMISO, FECHA_EMPRESA_COMPROMISO) VALUES ("&vid&", 1, 0, GETDATE())"
	
	conn.execute (queryBloqueo)

	set rsCorreo = conn.execute ("select EMAIL,EMAIL_CONTA,RUT_COORDINACION,RUT_CONTA,R_SOCIAL,RUT from EMPRESAS where ID_EMPRESA = '"&vid&"'")
	
	Set iMsg = CreateObject("CDO.Message")
	Set iConf = CreateObject("CDO.Configuration")
	Set Flds = iConf.Fields
	
	' send one copy with Google SMTP server (with autentication)
	schema = "http://schemas.microsoft.com/cdo/configuration/"
	Flds.Item(schema & "sendusing") = 2
	Flds.Item(schema & "smtpserver") = "smtp.gmail.com" 
	Flds.Item(schema & "smtpserverport") = 465
	Flds.Item(schema & "smtpauthenticate") = 1
	Flds.Item(schema & "sendusername") = "notificaciones@otecmutual.cl"
	Flds.Item(schema & "sendpassword") =  "admin_2019"
	Flds.Item(schema & "smtpusessl") = 1
	Flds.Update
	
	'creamos el nombre del archivo 
	archivo= Server.MapPath("mail.html")
	'conectamos con el FSO 
	set confile = createObject("scripting.filesystemobject") 
	'volvemos a abrir el fichero para lectura 
	set fich = confile.OpenTextFile(archivo) 
	
	'leemos el contenido del fichero 
	texto_fichero = fich.readAll() 
	'cerramos el fichero 
	fich.close()
	texto_fichero = Replace(texto_fichero,"<!--#empresa-->",rsCorreo("R_SOCIAL"))
	texto_fichero = Replace(texto_fichero,"<!--#rut_empresa-->",rsCorreo("RUT")) 
	texto_fichero = Replace(texto_fichero,"<!--#usuario-->",rsCorreo("EMAIL"))
	texto_fichero = Replace(texto_fichero,"<!--#pass-->",my_num)
	'response.End()
	
	With iMsg
	.To = rsCorreo("EMAIL")
	.From = "Inscripción de Empresa <notificaciones@otecmutual.cl>"
	.Subject = "Inscripción de Empresa"
	.HTMLBody = texto_fichero
	.Sender = "Mutual Capacitación"
	.Organization = "Mutual Capacitación"
	Set .Configuration = iConf
	SendEmailGmail = .Send
	End With
	
	set iMsg = nothing
	set iConf = nothing
	set Flds = nothing
	
	if(Request("tipoContacto")="0")then
	Set iMsg = CreateObject("CDO.Message")
	Set iConf = CreateObject("CDO.Configuration")
	Set Flds = iConf.Fields
	
	' send one copy with Google SMTP server (with autentication)
	schema = "http://schemas.microsoft.com/cdo/configuration/"
	Flds.Item(schema & "sendusing") = 2
	Flds.Item(schema & "smtpserver") = "smtp.gmail.com" 
	Flds.Item(schema & "smtpserverport") = 465
	Flds.Item(schema & "smtpauthenticate") = 1
	Flds.Item(schema & "sendusername") = "notificaciones@otecmutual.cl"
	Flds.Item(schema & "sendpassword") =  "admin_2019"
	Flds.Item(schema & "smtpusessl") = 1
	Flds.Update
	
	'creamos el nombre del archivo 
	archivo= Server.MapPath("mail.html")
	'conectamos con el FSO 
	set confile = createObject("scripting.filesystemobject") 
	'volvemos a abrir el fichero para lectura 
	set fich = confile.OpenTextFile(archivo) 
	
	'leemos el contenido del fichero 
	texto_fichero = fich.readAll() 
	'cerramos el fichero 
	fich.close() 
	texto_fichero = Replace(texto_fichero,"<!--#empresa-->",rsCorreo("R_SOCIAL")) 
	texto_fichero = Replace(texto_fichero,"<!--#rut_empresa-->",rsCorreo("RUT")) 
	texto_fichero = Replace(texto_fichero,"<!--#usuario-->",rsCorreo("EMAIL_CONTA"))
	texto_fichero = Replace(texto_fichero,"<!--#pass-->",my_num2)
	'response.End()
	
	With iMsg
	.To = rsCorreo("EMAIL_CONTA")
	.From = "Inscripción de Empresa <notificaciones@otecmutual.cl>"
	.Subject = "Inscripción de Empresa"
	.HTMLBody = texto_fichero
	.Sender = "Mutual Capacitación"
	.Organization = "Mutual Capacitación"
	Set .Configuration = iConf
	SendEmailGmail = .Send
	End With
	
	set iMsg = nothing
	set iConf = nothing
	set Flds = nothing
	end if

else
    set rsCorreo = conn.execute ("select EMAIL,R_SOCIAL,RUT from EMPRESAS where ID_EMPRESA = '"&vid&"'")
	
	Set iMsg = CreateObject("CDO.Message")
	Set iConf = CreateObject("CDO.Configuration")
	Set Flds = iConf.Fields
	
	' send one copy with Google SMTP server (with autentication)
	schema = "http://schemas.microsoft.com/cdo/configuration/"
	Flds.Item(schema & "sendusing") = 2
	Flds.Item(schema & "smtpserver") = "smtp.gmail.com" 
	Flds.Item(schema & "smtpserverport") = 465
	Flds.Item(schema & "smtpauthenticate") = 1
	Flds.Item(schema & "sendusername") = "notificaciones@otecmutual.cl"
	Flds.Item(schema & "sendpassword") =  "admin_2019"
	Flds.Item(schema & "smtpusessl") = 1
	Flds.Update
	
	'creamos el nombre del archivo 
	archivo= Server.MapPath("mailrechazo.html")
	'conectamos con el FSO 
	set confile = createObject("scripting.filesystemobject") 
	'volvemos a abrir el fichero para lectura 
	set fich = confile.OpenTextFile(archivo) 
	

	'leemos el contenido del fichero 
	texto_fichero = fich.readAll() 
	'cerramos el fichero 
	fich.close()
	
	texto_fichero = Replace(texto_fichero,"<!--#empresa-->",rsCorreo("R_SOCIAL"))
	texto_fichero = Replace(texto_fichero,"<!--#rut-->",replace(FormatNumber(mid(rsCorreo("RUT"), 1,len(rsCorreo("RUT"))-2),0)&mid(rsCorreo("RUT"), len(rsCorreo("RUT"))-1,len(rsCorreo("RUT"))),",","."))
	'response.End()
	
	With iMsg
	.To = rsCorreo("EMAIL")
	.From = "Inscripción de Empresa <notificaciones@otecmutual.cl>"
	.Subject = "Inscripción de Empresa"
	.HTMLBody = texto_fichero
	.Sender = "Mutual Capacitación"
	.Organization = "Mutual Capacitación"
	Set .Configuration = iConf
	SendEmailGmail = .Send
	End With
	
	set iMsg = nothing
	set iConf = nothing
	set Flds = nothing
end if

if err.number <> 0 then
	Response.Write("<commit>false</commit>")
else
	Response.Write("<commit>true</commit>")
end if
Response.Write("</modificar>") 
%>